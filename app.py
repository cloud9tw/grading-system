import os
import json
import datetime
import logging
import uuid
from google.cloud import bigquery
from google.oauth2 import service_account
from flask import Flask, redirect, url_for, session, render_template, request, jsonify, Response, flash
from authlib.integrations.flask_client import OAuth
import gspread
from dotenv import load_dotenv
import threading
from ceep_scraper import scrape_ceep_all_forms
from ceep_archiver import archive_to_sheets
from sync_to_bq import sync_all as sync_to_bq_all
from ai_handler import generate_ilp_chatgpt
from privacy_utils import get_code, decode_name
import pandas as pd
import io
from flask import send_file

load_dotenv()
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "super-secret-key-12345")

from werkzeug.middleware.proxy_fix import ProxyFix
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

# 強制 Session 設定以修復 OAuth MismatchingStateError
app.config.update(
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    PERMANENT_SESSION_LIFETIME=datetime.timedelta(days=1)
)

# OAuth Setup
oauth = OAuth(app)
google = oauth.register(
    name='google',
    client_id=os.getenv("GOOGLE_CLIENT_ID"),
    client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
    access_token_url='https://oauth2.googleapis.com/token',
    access_token_params=None,
    authorize_url='https://accounts.google.com/o/oauth2/auth',
    authorize_params=None,
    api_base_url='https://www.googleapis.com/oauth2/v1/',
    userinfo_endpoint='https://openidconnect.googleapis.com/v1/userinfo',
    client_kwargs={'scope': 'openid email profile'},
    server_metadata_url='https://accounts.google.com/.well-known/openid-configuration'
)

# 全域快取 (Global Cache) 減少 Sheets API 呼叫
GLOBAL_CACHE = {
    'students': {'data': None, 'time': None},
    'teachers': {'data': None, 'time': None},
    'stations': {'data': None, 'time': None},
    'epa_requirements': {'data': None, 'time': None},
    'exam_rooms': {'data': None, 'time': None},
    'scoring_config': {'data': None, 'time': None}
}
CACHE_TTL = 600 # 10 分鐘

def get_cached_data(key, fetch_func):
    now = datetime.datetime.now()
    cache = GLOBAL_CACHE.get(key)
    if cache and cache['data'] and cache['time'] and (now - cache['time']).total_seconds() < CACHE_TTL:
        return cache['data']
    
    # 執行抓取
    data = fetch_func()
    GLOBAL_CACHE[key] = {'data': data, 'time': now}
    return data

def safe_get_all_records(worksheet):
    data = worksheet.get_all_values()
    if not data:
        return []
    headers = data[0]
    records = []
    for row in data[1:]:
        record = {}
        for i, h in enumerate(headers):
            key = str(h).strip()
            if key:
                val = row[i] if i < len(row) else ''
                record[key] = val
        records.append(record)
    return records

from credentials_utils import get_bq_client, get_gspread_client

@app.route('/api/health')
def health_check():
    """系統健康檢查 API：診斷環境變數、Google Sheets 與 BigQuery 連線"""
    status = {
        'status': 'healthy',
        'timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'services': {
            'environment': 'ok',
            'google_sheets': 'pending',
            'bigquery': 'pending'
        }
    }
    
    # 1. 檢查關鍵環境變數
    required = ["GOOGLE_SHEET_ID", "SENDER_EMAIL"]
    missing = [e for e in required if not os.getenv(e)]
    if missing:
        status['status'] = 'degraded'
        status['services']['environment'] = f'Missing: {", ".join(missing)}'

    # 2. 檢查 Google Sheets
    try:
        gc = get_gspread_client()
        sid = os.getenv("GOOGLE_SHEET_ID")
        gc.open_by_key(sid)
        status['services']['google_sheets'] = 'ok'
    except Exception as e:
        status['status'] = 'unhealthy'
        status['services']['google_sheets'] = f"Connection Failed: {str(e)}"

    # 3. 檢查 BigQuery
    try:
        client, _ = get_bq_client()
        if client:
            status['services']['bigquery'] = 'ok'
        else:
            status['services']['bigquery'] = 'Initialization Failed'
    except Exception as e:
        if status['status'] == 'healthy': status['status'] = 'degraded'
        status['services']['bigquery'] = str(e)

    code = 200 if status['status'] != 'unhealthy' else 503
    return jsonify(status), code

@app.route('/')
def index():
    user = session.get('user')
    student_id = request.args.get('student_id')
    
    if student_id and not user:
        session['next_url'] = f"/?student_id={student_id}"
        
    if user:
        current_role = session.get('current_role', 'teacher')
        roles = session.get('roles', [])
        if current_role == 'student':
            import urllib.parse
            # 確保使用正式 student_info，不使用 preview_student_info
            student_info = session.get('student_info', {})
            
            # student_info 為空時（如管理員無對應學員資料），退回教師模式
            if not student_info or not student_info.get('id'):
                session['current_role'] = 'teacher'
                session.pop('preview_student_info', None)
                session.pop('is_preview_mode', None)
                return render_template('dashboard.html', user=user, roles=roles)
            
            s_id = student_info.get('id', '')
            root_url = request.url_root
            target_url = f"{root_url}attendance?student_id={s_id}"
            qr_img_url = f"https://api.qrserver.com/v1/create-qr-code/?size=250x250&data={urllib.parse.quote(target_url)}"
            # 切換至學生模式時，同步清除管理員預覽殘留狀態
            session.pop('preview_student_info', None)
            session.pop('is_preview_mode', None)
            return render_template('student_dashboard.html', user=user, student_info=student_info, qr_img_url=qr_img_url, roles=roles)
        else:
            return render_template('dashboard.html', user=user, roles=roles)
            
    return render_template('login.html')

import smtplib
from email.message import EmailMessage
import threading

def send_attendance_alert_email(student_name, teacher_name, sub_room, action, timestamp, diff_minutes):
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    notify_emails = os.getenv("NOTIFY_EMAILS")
    
    if not sender_email or not sender_password or not notify_emails:
        print("未設定 Email 環境變數，跳過發信機制。")
        return
        
    notify_list = [e.strip() for e in notify_emails.split(',')]
    msg = EmailMessage()
    
    if action == 'check_in':
        subject = f"⚠️ [遲到警報] 學員出勤異常通知"
        body = f"系統偵測到以下學員遭遇「遲到」異常：\n\n【學員姓名】：{student_name}\n【簽到時間】：{timestamp} (遲到約 {diff_minutes} 分鐘)\n【檢查室別】：{sub_room}\n【紀錄教師】：{teacher_name}\n\n※ 此為系統自動發送之信件，請勿回覆。"
    else:
        subject = f"⚠️ [早退警報] 學員出勤異常通知"
        body = f"系統偵測到以下學員遭遇「早退」異常：\n\n【學員姓名】：{student_name}\n【簽退時間】：{timestamp} (早退約 {diff_minutes} 分鐘)\n【檢查室別】：{sub_room}\n【紀錄教師】：{teacher_name}\n\n※ 此為系統自動發送之信件，請勿回覆。"
        
    msg.set_content(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = ", ".join(notify_list)
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
            print(f"成功發送出勤警報給 {notify_emails}")
    except Exception as e:
        print(f"發送出勤警報失敗: {e}")

def send_access_request_email(teacher_name, teacher_email):
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    admin_emails = os.getenv("NOTIFY_EMAILS")
    
    if not sender_email or not sender_password or not admin_emails:
        print("未設定 Email 環境變數，無法發送申請通知。")
        return
        
    msg = EmailMessage()
    subject = f"🔔 [權限申請] 新教師登入權限請求 - {teacher_name}"
    body = f"【學員評分系統 - 權限申請通知】\n\n有新教師申請使用個人 Google 帳號登入系統：\n\n"
    body += f"● 教師姓名：{teacher_name}\n"
    body += f"● Gmail 帳號：{teacher_email}\n\n"
    body += f"請管理員核對身分後，將此 Email 新增至「教師名單」試算表中以開通權限。\n\n"
    body += f"※ 此為系統自動發送之申請信件，請勿回覆。"
    
    msg.set_content(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = admin_emails
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
            print(f"✅ 成功將 {teacher_name} 的權限申請發送給管理員")
    except Exception as e:
        print(f"❌ 發送權限申請失敗: {e}")

def send_feedback_anomaly_email(admin_emails, data, triggers):
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    
    if not sender_email or not sender_password or not admin_emails:
        print("未設定 Email 環境變數或無管理員名單，跳過回饋警報機制。")
        return
        
    student_name = data.get('student_name', 'Unknown')
    teacher_name = data.get('teacher', 'Unknown')
    station = data.get('station', 'Unknown')
    suggestion = data.get('suggestion', '(無文字建議)')
    
    msg = EmailMessage()
    subject = f"🔴 [教學品質警報] {teacher_name} 的臨床教學收到負面回饋"
    
    trigger_text = ", ".join(triggers)
    body = f"【教學品質警報系統】\n\n系統偵測到以下臨床教學回饋符合負面預警標準，請相關人員儘速查證：\n\n"
    body += f"● 觸發原因：{trigger_text}\n"
    body += f"● 學員姓名：{student_name}\n"
    body += f"● 受評教師：{teacher_name}\n"
    body += f"● 實習站別：{station}\n\n"
    body += f"【學員之整體建議】：\n{suggestion}\n\n"
    body += f"※ 詳細各項評分請至 Google Sheets 查詢：https://docs.google.com/spreadsheets/d/112l_e3WKbIkFYj58nv8LRTYEvfyDpXMh-NcSe98T07w/edit\n\n"
    body += f"※ 此為系統自動發送之警示，請勿回覆。"
    
    msg.set_content(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = ", ".join(admin_emails)
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
            msg_ok = f"✅ 教學品質警報已成功發送給管理員群組: {admin_emails}\n"
            print(msg_ok)
            with open('feedback_debug.log', 'a', encoding='utf-8') as f:
                f.write(f"  {msg_ok}\n")
    except Exception as e:
        msg_err = f"❌ 發送教學品質警報失敗: {e}\n"
        print(msg_err)
        with open('feedback_debug.log', 'a', encoding='utf-8') as f:
            f.write(f"  {msg_err}\n")

@app.route('/login')
def login():
    # ===== 測試模式：bypass OAuth =====
    if os.getenv('TEST_MODE', '').lower() == 'true':
        test_role = request.args.get('role', 'teacher')  # ?role=student 或 ?role=teacher
        test_user = {
            'name': '測試教師' if test_role == 'teacher' else '測試學員',
            'email': 'test-teacher@test.com' if test_role == 'teacher' else 'test-student@test.com',
            'picture': 'https://ui-avatars.com/api/?name=Test+User&background=random'
        }
        session['user'] = test_user
        session['is_admin'] = (test_role == 'teacher')

        if test_role == 'student':
            session['roles'] = ['student']
            session['current_role'] = 'student'
            session['student_info'] = {
                'id': 'TEST001',
                'name': '測試學員',
                'type': '住院醫師',
                'gender': '男'
            }
        else:
            session['roles'] = ['teacher']
            session['current_role'] = 'teacher'

        next_url = session.pop('next_url', '/')
        return redirect(next_url)
    # ===== 正常 OAuth 流程 =====
    redirect_uri = url_for('authorize', _external=True)
    return google.authorize_redirect(redirect_uri)

@app.route('/api/request_access', methods=['POST'])
def request_access():
    teacher_name = request.form.get('teacher_name', '').strip()
    teacher_email = request.form.get('teacher_email', '').strip()
    
    if not teacher_name or not teacher_email:
        return "請提供姓名與 Email", 400
        
    # 非同步發送郵件
    threading.Thread(target=send_access_request_email, args=(teacher_name, teacher_email)).start()
    
    return redirect('/?request_sent=1')

@app.route('/authorize')
def authorize():
    token = google.authorize_access_token()
    user_info = token.get('userinfo')
    if not user_info:
        user_info = google.userinfo()
    session['user'] = user_info
    
    user_email = user_info.get('email', '').strip().lower()
    roles = []
    student_info = None
    
    try:
        gc = get_gspread_client()
        if gc:
            sheet_id = os.getenv("GOOGLE_SHEET_ID")
            doc = gc.open_by_key(sheet_id)
            
            # Check Teacher
            sheet_teachers = doc.worksheet('教師名單')
            teachers_records = safe_get_all_records(sheet_teachers)
            is_admin = False
            for t in teachers_records:
                t_email = str(t.get('教師_Email', '')).strip().lower()
                if t_email == user_email:
                    roles.append('teacher')
                    # Detect Admin Privilege
                    if str(t.get('管理員權限', '')).strip().lower() == 'admin':
                        is_admin = True
                    break
            session['is_admin'] = is_admin
                    
            # Check Student
            sheet_students = doc.worksheet('學員名單')
            students_records = safe_get_all_records(sheet_students)
            for s in students_records:
                s_email = str(s.get('Email', '')).strip().lower()
                if s_email == user_email:
                    roles.append('student')
                    student_info = {
                        'id': str(s.get('學生ID', '')).strip(),
                        'name': str(s.get('姓名', '')),
                        'type': str(s.get('學員類別', '')),
                        'gender': str(s.get('性別', ''))
                    }
                    break
    except Exception as e:
        print("Role check error:", e)
        
    session['roles'] = roles
    if student_info:
        session['student_info'] = student_info
        
    next_url = session.pop('next_url', '/')
    
    # Determine current role
    default_role = 'teacher'
    if 'student' in roles:
        default_role = 'student'
        
    # If they scanned a QR code string, force teacher
    if 'student_id=' in next_url and 'teacher' in roles:
        default_role = 'teacher'
        
    session['current_role'] = default_role
    
    return redirect(next_url)

@app.route('/switch_role', methods=['POST'])
def switch_role():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    roles = session.get('roles', [])
    current = session.get('current_role', 'teacher')
    
    if 'student' in roles and 'teacher' in roles:
        session['current_role'] = 'teacher' if current == 'student' else 'student'
        return jsonify({'success': True})
    return jsonify({'success': False, 'error': '無切換權限'})

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect('/')

@app.route('/attendance')
def attendance():
    user = session.get('user')
    if not user:
        # 保存帶有 student_id 的完整 URL，讓登入後可以自動帶回
        next_url = request.url
        # 避免 next_url 包含 host，只保留 path+query
        from urllib.parse import urlparse
        parsed = urlparse(next_url)
        session['next_url'] = parsed.path + ('?' + parsed.query if parsed.query else '')
        return redirect('/login')
    return render_template('attendance.html', user=user, roles=session.get('roles', []))

@app.route('/feedback')
def feedback_page():
    user = session.get('user')
    if not user:
        session['next_url'] = '/feedback'
        return redirect('/login')
    # 僅限學員身份使用
    if session.get('current_role') != 'student':
        return render_template('error.html', message='教學回饋表僅限學員身份使用。若您同時具有教師與學員身份，請先切換至「學員介面」再點擊。'), 403
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        # 取得教師名單
        teachers_records = safe_get_all_records(doc.worksheet('教師名單'))
        teachers = [str(r.get('教師姓名', '')).strip() for r in teachers_records if str(r.get('教師姓名', '')).strip()]
        # 取得檢查室清單（與簽到退相同資料源）
        room_sheet = doc.worksheet('檢查室清單')
        all_rows = room_sheet.get_all_values()
        departments = []
        for row in all_rows:
            dept = str(row[0]).strip() if row else ''
            if not dept:
                continue
            rooms = [str(c).strip() for c in row[1:] if str(c).strip()]
            departments.append({'dept': dept, 'rooms': rooms})
        # 從學員名單中，依 Email 比對取得學生姓名
        user_email = user.get('email', '').strip().lower()
        students_records = safe_get_all_records(doc.worksheet('學員名單'))
        student_name = user.get('name', '')  # 預設為 Google 帳號姓名
        for rec in students_records:
            rec_email = str(rec.get('Email', '')).strip().lower()
            if rec_email and rec_email == user_email:
                matched = str(rec.get('姓名', '')).strip()
                if matched:
                    student_name = matched
                break
    except Exception as e:
        teachers = []
        departments = []
        student_name = user.get('name', '')
    import json
    return render_template('feedback.html', user=user, roles=session.get('roles', []),
                           student_name=student_name,
                           teachers=teachers, departments_json=json.dumps(departments, ensure_ascii=False))

@app.route('/api/submit_feedback', methods=['POST'])
def submit_feedback():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    try:
        data = request.json
        gc = get_gspread_client()
        FEEDBACK_SHEET_ID = '112l_e3WKbIkFYj58nv8LRTYEvfyDpXMh-NcSe98T07w'
        doc = gc.open_by_key(FEEDBACK_SHEET_ID)
        sheet = doc.worksheet('表單回應')

        now_dt = datetime.datetime.utcnow() + datetime.timedelta(hours=8)
        timestamp = now_dt.strftime('%Y/%m/%d 上午 %I:%M:%S')

        def to_int(v):
            try: return int(v)
            except: return ''

        # 前端傳來的各類分組
        ability  = data.get('ability', {})    # 整體能力 H-K (q1–q4)
        teaching = data.get('teaching', {})   # 教學活動 L-Q (q1–q6)
        holistic = data.get('holistic', {})   # 全人醫療 R-U (q1–q4)
        knowledge= data.get('knowledge', {})  # 醫學知識 V-X (q1–q3)
        skills   = data.get('skills', {})     # 教學技巧 Y-AA (q1–q3)

        # 欄次： A B C D E F G | H I J K | L M N O P Q | R S T U | V W X | Y Z AA | AB AC AD AE AF AG | AH | AI
        row = [
            '',                                                                # A (斷行編號，留空)
            timestamp,                                                         # B 時間戳記
            data.get('student_name', user.get('name', '')),                    # C 學生姓名
            user.get('email', ''),                                             # D 電子郵件
            data.get('teacher', ''),                                           # E 教師名稱
            data.get('other_teacher', ''),                                     # F 未登錄教師
            data.get('station', ''),                                           # G 臨床實習站別
            # H–K: 整體能力 (老師能提升, 事先了解, 充分學習, 尊重學生)
            to_int(ability.get('q1')), to_int(ability.get('q2')),
            to_int(ability.get('q3')), to_int(ability.get('q4')),
            # L–Q: 教學活動 (目標, 臨實, 示範, 問題導向, 鼓勵提問, 筆記)
            to_int(teaching.get('q1')), to_int(teaching.get('q2')),
            to_int(teaching.get('q3')), to_int(teaching.get('q4')),
            to_int(teaching.get('q5')), to_int(teaching.get('q6')),
            # R–U: 全人醫療 (照護病人, 醫療溝通, 倫理社會, 團隊)
            to_int(holistic.get('q1')), to_int(holistic.get('q2')),
            to_int(holistic.get('q3')), to_int(holistic.get('q4')),
            # V–X: 醫學知識 (專業知識, 實證, 基礎+臨床)
            to_int(knowledge.get('q1')), to_int(knowledge.get('q2')),
            to_int(knowledge.get('q3')),
            # Y–AA: 教學技巧 (自我學習, 回饋評核, 排定活動)
            to_int(skills.get('q1')), to_int(skills.get('q2')),
            to_int(skills.get('q3')),
            # AB–AF: 各類平均 (公式自動計算，留空)
            '', '', '', '', '',
            # AG: 平均 (公式，留空)
            '',
            # AH: 對教師的整體建議
            data.get('suggestion', ''),
            # AI: 醫院(公式，留空)
            ''
        ]
        sheet.append_row(row, table_range="A1")

        # --- 智慧預警檢測 (新功能) ---
        def process_feedback_alert(feedback_data):
            try:
                gc_alert = get_gspread_client()
                sheet_main_id = os.getenv("GOOGLE_SHEET_ID")
                doc_alert = gc_alert.open_by_key(sheet_main_id)
                
                # 1. 獲取所有管理員 Email
                sheet_teachers = doc_alert.worksheet('教師名單')
                teacher_recs = safe_get_all_records(sheet_teachers)
                admin_emails = []
                for tr in teacher_recs:
                    # 檢查 E 欄 (管理員權限) 是否標註為 Admin
                    if str(tr.get('管理員權限', '')).strip().lower() == 'admin':
                        email = str(tr.get('教師_Email', '')).strip().lower()
                        if email and '@' in email:
                            admin_emails.append(email)
                
                # 2. 獲取自定義負面關鍵字
                try:
                    sheet_settings = doc_alert.worksheet('系統設定')
                    kw_recs = sheet_settings.col_values(1) # 負面關鍵字在第一欄
                    custom_keywords = [str(k).strip() for k in kw_recs[1:] if str(k).strip()]
                except:
                    custom_keywords = []
                
                # 3. 執行檢修邏輯
                alert_triggers = []
                all_scores = []
                
                # 遍歷所有量化組別查低分 (q1, q2...)
                for group in ['ability', 'teaching', 'holistic', 'knowledge', 'skills']:
                    g_data = feedback_data.get(group, {})
                    for q, val in g_data.items():
                        try:
                            v = int(val)
                            all_scores.append(v)
                            if v < 4:
                                alert_triggers.append(f"單項低分({v}分)")
                        except: pass
                
                # 整體平均檢查
                if all_scores:
                    avg = sum(all_scores) / len(all_scores)
                    if avg < 4.0:
                        alert_triggers.append(f"整體平均未達優(不足4分)")
                
                # 關鍵字智慧判定 (針對對教師建議)
                suggestion_text = str(feedback_data.get('suggestion', ''))
                if suggestion_text:
                    for kw in custom_keywords:
                        if kw in suggestion_text:
                            alert_triggers.append(f"命中負面詞: {kw}")
                            break
                
                # 4. 記錄偵錯資訊
                log_msg = f"[{timestamp}] 檢測開始: 學員={feedback_data.get('student_name')}, 教師={feedback_data.get('teacher')}\n"
                log_msg += f"  - 管理員名單: {len(admin_emails)} 人\n"
                log_msg += f"  - 自定義關鍵字: {len(custom_keywords)} 個\n"
                log_msg += f"  - 觸發項目: {alert_triggers}\n"
                
                with open('feedback_debug.log', 'a', encoding='utf-8') as f:
                    f.write(log_msg)
                
                # 如果符合任一預警條件且有名單，則寄信
                if alert_triggers and admin_emails:
                    send_feedback_anomaly_email(list(set(admin_emails)), feedback_data, list(set(alert_triggers)))
                    with open('feedback_debug.log', 'a', encoding='utf-8') as f:
                        f.write(f"  ✅ 已嘗試發送信件至 {len(admin_emails)} 位管理員\n")
                    
            except Exception as ex:
                err_msg = f"智慧預警背景處理異常: {ex}\n"
                print(err_msg)
                with open('feedback_debug.log', 'a', encoding='utf-8') as f:
                    f.write(f"  ❌ 錯誤: {err_msg}\n")

        # 啟動背景執行緒處理，不影響 API 回傳速度
        threading.Thread(target=process_feedback_alert, args=(data,)).start()
        
        # BQ Double-Write
        try:
            bq_client, project_id = get_bq_client()
            if bq_client:
                dataset_id = f"{project_id}.grading_data"
                fb_table_id = f"{dataset_id}.feedback_logs"
                ts_iso = now_dt.isoformat()
                fb_insert = [{
                    "timestamp": ts_iso,
                    "email": user.get('email', ''),
                    "student_name": data.get('student_name', ''),
                    "role": data.get('role', ''),
                    "teacher": data.get('teacher', ''),
                    "co_teacher": data.get('co_teacher', ''),
                    "department": data.get('department', ''),
                    "is_retake": data.get('is_retake', ''),
                    "score": data.get('score', ''),
                    "suggestions": data.get('suggestions', '')
                }]
                errors = bq_client.insert_rows_json(fb_table_id, fb_insert)
                if errors: print("BQ Feedback Insert Error:", errors)
        except Exception as e:
            print("BQ Insert Error (Feedback):", e)
            
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/qrcodes')
def qrcodes():
    user = session.get('user')
    if not user:
        session['next_url'] = '/qrcodes'
        return redirect('/login')
        
    try:
        gc = get_gspread_client()
        if not gc:
            return "Google Sheets backend not configured.", 500
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        sheet_students = doc.worksheet('學員名單')
        students_records = safe_get_all_records(sheet_students)
        students = [{'id': str(rec.get('學生ID', '')), 'name': str(rec.get('姓名', ''))} for rec in students_records if rec.get('姓名')]
        
        # request.url_root returns something like 'https://example.com/'
        import urllib.parse
        root_url = request.url_root
        for s in students:
            target_url = f"{root_url}attendance?student_id={s['id']}"
            s['qr_img_url'] = f"https://api.qrserver.com/v1/create-qr-code/?size=250x250&data={urllib.parse.quote(target_url)}"
            
        return render_template('qrcodes.html', user=user, students=students)
    except Exception as e:
        return f"Error loading student roster: {str(e)}", 500

@app.route('/api/attendance_config', methods=['GET'])
def get_attendance_config():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    try:
        gc = get_gspread_client()
        if not gc:
            return jsonify({'success': False, 'error': 'Google Sheets backend not configured.'}), 500
            
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        # 取得學員名單
        sheet_students = doc.worksheet('學員名單')
        students_records = safe_get_all_records(sheet_students)
        students = [{'id': str(rec.get('學生ID', '')), 'name': str(rec.get('姓名', ''))} for rec in students_records if rec.get('姓名')]
        
        # 取得教師名單
        try:
            sheet_teachers = doc.worksheet('教師名單')
            teachers_records = safe_get_all_records(sheet_teachers)
            teachers = [{'name': str(rec.get('教師姓名', '')).strip()} for rec in teachers_records if str(rec.get('教師姓名', '')).strip()]
        except Exception:
            teachers = []
        
        # 取得檢查室大項與次項目
        sheet_rooms = doc.worksheet('檢查室清單')
        rooms_data = sheet_rooms.get_all_values()
        departments = []
        for row in rooms_data:
            if not row or not str(row[0]).strip():
                continue
            main_dept = str(row[0]).strip()
            sub_rooms = [str(col).strip() for col in row[1:] if str(col).strip()]
            departments.append({
                'main': main_dept,
                'sub_rooms': sub_rooms
            })
            
        return jsonify({
            'success': True,
            'students': students,
            'teachers': teachers,
            'departments': departments
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/submit_attendance', methods=['POST'])
def submit_attendance():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    data = request.json
    try:
        gc = get_gspread_client()
        if not gc:
            return jsonify({'success': False, 'error': 'Google Sheets backend not configured.'}), 500
            
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        sheet = doc.worksheet('上下班打卡記錄')
        
        import datetime
        timestamp = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
        teacher_email = user.get('email', '')
        teacher_name = teacher_email # default
        
        # 查詢教師名單中的姓名
        try:
            sheet_teachers = doc.worksheet('教師名單')
            teachers_data = safe_get_all_records(sheet_teachers)
            for rec in teachers_data:
                t_email = str(rec.get('教師_Email', '')).strip().lower()
                if t_email == teacher_email.lower():
                    n = str(rec.get('教師姓名', '')).strip()
                    if n:
                        teacher_name = n
                    break
        except:
            pass
            
        student_name = data.get('student_name', '').split(' (')[0].strip()
        sub_room = data.get('sub_room', '')
        action = data.get('action', '') # 'check_in' or 'check_out'
        co_teacher = data.get('co_teacher', '') # Optional Co-teacher
        
        if not student_name or not sub_room or not action:
            return jsonify({'success': False, 'error': '資料不齊全'}), 400
            
        now_dt = datetime.datetime.utcnow() + datetime.timedelta(hours=8)
        cur_time = now_dt.time()
        
        checkin_std = datetime.time(8, 40, 0)
        checkout_std = datetime.time(17, 0, 0)
        
        alert_needed = False
        time_diff = 0
        
        if action == 'check_in':
            if cur_time > checkin_std:
                dt_std = datetime.datetime.combine(now_dt.date(), checkin_std)
                diff_minutes = int((now_dt - dt_std).total_seconds() / 60)
                if diff_minutes > 0:
                    alert_needed = True
                    time_diff = diff_minutes
                    
            row = [student_name, teacher_name, co_teacher, sub_room, timestamp, '']
            sheet.append_row(row, table_range="A1")
            
            try:
                bq_client, project_id = get_bq_client()
                if bq_client:
                    t_iso = now_dt.isoformat()
                    bq_client.insert_rows_json(f"{project_id}.grading_data.attendance_events", [{
                        "student_name": student_name, "teacher_name": teacher_name, "co_teacher": co_teacher,
                        "sub_room": sub_room, "event_type": "CHECK_IN", "event_time": t_iso
                    }])
            except Exception as e:
                print("BQ Attendance Check-in Error:", e)
                

            if alert_needed:
                threading.Thread(target=send_attendance_alert_email, args=(student_name, teacher_name, sub_room, action, timestamp, time_diff)).start()
                return jsonify({'success': True, 'msg': f'✅ 簽到成功！(⚠️ 系統偵測到遲到 {time_diff} 分鐘，已發信通報)'})
            return jsonify({'success': True, 'msg': '✅ 簽到成功！'})
            
        elif action == 'check_out':
            if cur_time < checkout_std:
                dt_std = datetime.datetime.combine(now_dt.date(), checkout_std)
                diff_minutes = int((dt_std - now_dt).total_seconds() / 60)
                if diff_minutes > 0:
                    alert_needed = True
                    time_diff = diff_minutes
                    
            all_vals = sheet.get_all_values()
            
            # 從最底下往上找該學員該次簽到
            found_idx = -1
            for i in range(len(all_vals)-1, 0, -1):
                r = all_vals[i]
                if len(r) >= 4 and str(r[0]).strip() == student_name and str(r[3]).strip() == sub_room:
                    # 如果簽退欄還沒填過 (第六欄 index 5)
                    if len(r) < 6 or not str(r[5]).strip():
                        found_idx = i
                        break
            
            if found_idx != -1:
                row_num = found_idx + 1
                sheet.update_cell(row_num, 6, timestamp)
                
                try:
                    bq_client, project_id = get_bq_client()
                    if bq_client:
                        t_iso = now_dt.isoformat()
                        bq_client.insert_rows_json(f"{project_id}.grading_data.attendance_events", [{
                            "student_name": student_name, "teacher_name": teacher_name, "co_teacher": co_teacher,
                            "sub_room": sub_room, "event_type": "CHECK_OUT", "event_time": t_iso
                        }])
                except Exception as e:
                    print("BQ Attendance Check-out Error:", e)

                if alert_needed:
                    threading.Thread(target=send_attendance_alert_email, args=(student_name, teacher_name, sub_room, action, timestamp, time_diff)).start()
                    return jsonify({'success': True, 'msg': f'📤 簽退成功！(⚠️ 系統偵測到早退 {time_diff} 分鐘，已通報)'})
                return jsonify({'success': True, 'msg': '📤 簽退成功！'})
            else:
                # 依指示：若檢查室不一致或找不到紀錄，強迫安插新的一列
                row = [student_name, teacher_name, co_teacher, sub_room, '', timestamp]
                sheet.append_row(row, table_range="A1")
                
                try:
                    bq_client, project_id = get_bq_client()
                    if bq_client:
                        t_iso = now_dt.isoformat()
                        bq_client.insert_rows_json(f"{project_id}.grading_data.attendance_events", [{
                            "student_name": student_name, "teacher_name": teacher_name, "co_teacher": co_teacher,
                            "sub_room": sub_room, "event_type": "CHECK_OUT", "event_time": t_iso
                        }])
                except Exception as e:
                    print("BQ Attendance Check-out Error:", e)

                if alert_needed:
                    threading.Thread(target=send_attendance_alert_email, args=(student_name, teacher_name, sub_room, action, timestamp, time_diff)).start()
                    return jsonify({'success': True, 'msg': f'⚠️ 無有效簽到紀錄強制簽退！(且早退 {time_diff} 分，已通報)'})
                return jsonify({'success': True, 'msg': '⚠️ 查無相對應的有效簽到紀錄，已新建獨立新列！'})
            
        else:
            return jsonify({'success': False, 'error': '未知的操作類型。'}), 400
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/cron/check_absent', methods=['GET'])
def check_absent():
    try:
        # 1. 初始化連線
        from credentials_utils import get_gspread_client
        gc = get_gspread_client()
        if not gc:
            return jsonify({'success': False, 'error': 'Google Sheets client initialization failed'}), 500
            
        now_dt = datetime.datetime.utcnow() + datetime.timedelta(hours=8)
        if now_dt.weekday() >= 5: # 週末不寄信
            return jsonify({'success': True, 'msg': 'Today is weekend, skip absent check.'})
            
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        today_str = now_dt.strftime("%Y-%m-%d")
        
        # 2. 檢查排除日期 (動態搜尋「排除日期」欄位)
        try:
            sheet_settings = doc.worksheet('系統設定')
            settings_vals = sheet_settings.get_all_values()
            if settings_vals:
                header = settings_vals[0]
                try:
                    # 尋找「排除日期」所在的標題欄位
                    date_col_idx = -1
                    for idx, val in enumerate(header):
                        if '排除日期' in str(val) or '國定假日' in str(val):
                            date_col_idx = idx
                            break
                    
                    if date_col_idx != -1:
                        excluded_dates = [str(r[date_col_idx]).strip() for r in settings_vals[1:] if len(r) > date_col_idx]
                        if today_str in excluded_dates:
                            return jsonify({'success': True, 'msg': f'Today ({today_str}) is an excluded date, skip absent check.'})
                except Exception as ex:
                    print(f"Search exclusion column error: {ex}")
        except Exception as e:
            print(f"Read excluded dates error: {e}")

        # 2. 取得所有學員
        sheet_students = doc.worksheet('學員名單')
        students_records = safe_get_all_records(sheet_students)
        all_students = [str(rec.get('姓名', '')).strip() for rec in students_records if str(rec.get('姓名', '')).strip()]
        if not all_students:
            return jsonify({'success': True, 'msg': 'No students in roster'})
        
        # 取得今日打卡
        sheet_records = doc.worksheet('上下班打卡記錄')
        all_vals = sheet_records.get_all_values()
        
        checked_in_students = set()
        for r in all_vals[1:]: # skip header
            # 簽到時間在 index 4 (0-indexed -> r[4])
            if len(r) > 4:
                checkin_time = str(r[4]).strip()
                if checkin_time.startswith(today_str):
                    student = str(r[0]).strip()
                    checked_in_students.add(student)
                    
        absent_students = [s for s in all_students if s not in checked_in_students]
        
        if not absent_students:
            return jsonify({'success': True, 'msg': 'All students have checked in today.'})
            
        # Send Email
        sender_email = os.getenv("SENDER_EMAIL")
        sender_password = os.getenv("SENDER_PASSWORD")
        notify_emails = os.getenv("NOTIFY_EMAILS")
        
        if sender_email and sender_password and notify_emails:
            import smtplib
            from email.message import EmailMessage
            msg = EmailMessage()
            subject = f"⚠️ [遲到警報] {today_str} 未簽到學員名單統整"
            
            absent_str = "\n".join([f"- {s}" for s in absent_students])
            body = f"系統偵測到以下學員今日 ({today_str}) 尚未完成簽到：\n\n{absent_str}\n\n※ 此為系統定時自動發送之信件，請勿回覆。"
            
            msg.set_content(body)
            msg['Subject'] = subject
            msg['From'] = sender_email
            msg['To'] = ", ".join([e.strip() for e in notify_emails.split(',')])
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(sender_email, sender_password)
                smtp.send_message(msg)
                
        return jsonify({'success': True, 'absent_count': len(absent_students), 'absent_students': absent_students})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/sync_ceep', methods=['GET', 'POST'])
def sync_ceep():
    user = session.get('user')
    if not user or not session.get('is_admin'):
        if request.is_json:
            return jsonify({'success': False, 'error': '僅限管理員執行'}), 403
        return redirect('/login')

    if request.method == 'GET':
        return render_template('sync_status.html')
    
    # 舊的 POST 方式保留，但建議改用串流
    return jsonify({'success': False, 'error': '請使用串流介面'})

@app.route('/api/sync_ceep_stream')
def sync_ceep_stream():
    """
    透過 Server-Sent Events (SSE) 串流傳輸同步進度。
    """
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return "Unauthorized", 403

    def generate_simple():
        import queue
        import threading
        import asyncio
        import json
        
        # 建立一個同步隊列
        msg_queue = queue.Queue()

        # 立即發送一個啟始訊息，確保串流通道已建立
        msg_queue.put("串流通道已建立，正在啟動背景同步執行緒...")

        async def callback(msg):
            msg_queue.put(msg)

        async def run_scraper():
            try:
                # 1. 第一階段：執行 Playwright 爬蟲
                msg_queue.put("正在進入非同步爬蟲核心...")
                data, summary = await scrape_ceep_all_forms(callback=callback)
                
                # 2. 第二階段：將結果存回 Sheets
                msg_queue.put("--- 正在將數據同步至 Google Sheets ---")
                for sheet_name, records in data.items():
                    archive_to_sheets(records, sheet_name=sheet_name)
                    msg_queue.put(f"✅ 已完成 {sheet_name} 同步")
                
                # 3. 第三階段：同步至 BigQuery (以 BQ 為核心)
                msg_queue.put("--- 正在將數據同步至 BigQuery (數據核心) ---")
                def bq_callback(m):
                    msg_queue.put(f"   [BQ] {m}")
                
                sync_success = sync_to_bq_all(callback=bq_callback)
                if sync_success:
                    msg_queue.put("✅ BigQuery 數據已同步完成")
                else:
                    msg_queue.put("⚠️ BigQuery 同步過程中有部分異常")

                msg_queue.put({"type": "summary", "data": summary})
                msg_queue.put({"type": "done", "msg": "所有同步任務已完成！"})
            except Exception as e:
                import traceback
                error_trace = traceback.format_exc()
                logging.error(f"Sync error traceback: {error_trace}")
                msg_queue.put({"type": "error", "msg": f"伺服器出錯: {str(e)}"})
            finally:
                msg_queue.put(None)

        def worker():
            logging.info("Starting sync worker thread...")
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(run_scraper())
            finally:
                loop.close()

        # 啟動背景執行緒
        t = threading.Thread(target=worker)
        t.daemon = True
        t.start()

        # 持續讀取隊列直到收到 None
        while True:
            try:
                msg = msg_queue.get(timeout=120) # 設定超時避免無限阻塞
                if msg is None:
                    break
                
                if isinstance(msg, dict):
                    yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
                else:
                    yield f"data: {json.dumps({'type': 'log', 'msg': msg}, ensure_ascii=False)}\n\n"
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'log', 'msg': '... 等待伺服器回應中 ...'}, ensure_ascii=False)}\n\n"

    response = Response(generate_simple(), mimetype='text/event-stream')
    # 關閉快取與緩衝，確保 SSE 即時性
    response.headers['Cache-Control'] = 'no-cache'
    response.headers['X-Accel-Buffering'] = 'no'
    response.headers['Connection'] = 'keep-alive'
    return response

# =====================================================================
# 📅 實習進度排程系統 (28 週日曆引擎)
# =====================================================================

def get_intern_year():
    """取得目前的實習學年度 (以 7 月為分界)"""
    now = datetime.datetime.now()
    return now.year if now.month >= 7 else now.year - 1

def get_internship_start(year=None):
    """
    取得實習開始日：指定年份 7 月的第一個星期一。
    例如 2026 年 -> 2026-07-06
    """
    if year is None:
        year = get_intern_year()
    d = datetime.date(year, 7, 1)
    # weekday() 0=Monday, 移動到下一個星期一
    days_ahead = (7 - d.weekday()) % 7
    return d + datetime.timedelta(days=days_ahead)

def get_current_intern_week(target_date=None):
    """
    計算今天（或指定日期）落在第幾週。
    - W1 = 7月第一個星期一起的 7 天
    - W27, W28 為「手動自選週」
    - 超出 28 週範圍回傳 None
    """
    if target_date is None:
        target_date = datetime.date.today()
    elif isinstance(target_date, datetime.datetime):
        target_date = target_date.date()

    start = get_internship_start()
    delta = (target_date - start).days
    if delta < 0:
        return None  # 尚未開始
    week = delta // 7 + 1
    if week > 28:
        return None  # 已超過實習期
    return week

SCHEDULE_SHEET_NAME = '學生進度排程'
MANUAL_WEEK_LABEL = '手動自選'

@app.route('/api/admin/schedule', methods=['GET'])
def get_admin_schedule():
    """讀取所有學生的 W1~W28 站別排程"""
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return jsonify({'success': False, 'error': '僅限管理員'}), 403
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)

        # 嘗試讀取排程表，若不存在則建立
        try:
            ws = doc.worksheet(SCHEDULE_SHEET_NAME)
        except Exception:
            # 建立新分頁，標頭為：學員姓名, 學員ID, W1, W2, ... W28
            headers = ['學員姓名', '學員ID'] + [f'W{i}' for i in range(1, 29)]
            ws = doc.add_worksheet(title=SCHEDULE_SHEET_NAME, rows=100, cols=len(headers))
            ws.append_row(headers)

        records = safe_get_all_records(ws)
        # 整理成 dict，key 為學員ID
        schedule_map = {}
        for rec in records:
            sid = str(rec.get('學員ID', '')).strip()
            if sid:
                weeks = {}
                for i in range(1, 29):
                    val = str(rec.get(f'W{i}', '')).strip()
                    # 解析以逗號分隔的複數站別
                    weeks[f'W{i}'] = [s.strip() for s in val.split(',') if s.strip()] if val else []
                schedule_map[sid] = {
                    'name': str(rec.get('學員姓名', '')).strip(),
                    'id': sid,
                    'weeks': weeks
                }

        # 取得小站清單：從「檢查室清單」讀取各部門下的子檢查室 (小站)
        try:
            room_sheet = doc.worksheet('檢查室清單')
            all_rows = room_sheet.get_all_values()
            stations = []
            seen = set()
            for row in all_rows:
                # 第一欄為大站名稱，第二欄起為各小站
                for cell in row[1:]:
                    val = str(cell).strip()
                    if val and val not in seen:
                        stations.append(val)
                        seen.add(val)
        except Exception:
            stations = []

        current_week = get_current_intern_week()
        intern_start = get_internship_start().isoformat()

        return jsonify({
            'success': True,
            'schedule': schedule_map,
            'stations': stations,
            'current_week': current_week,
            'intern_start': intern_start,
            'manual_weeks': [27, 28]  # 最後兩週為手動自選
        })
    except Exception as e:
        logging.error(f"get_admin_schedule error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/admin/schedule', methods=['POST'])
def save_admin_schedule():
    """儲存（覆蓋）指定學生的 W1~W28 排程"""
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return jsonify({'success': False, 'error': '僅限管理員'}), 403
    try:
        data = request.json
        # data 格式: { student_id: str, student_name: str, weeks: { W1: ["CT","MRI"], W2: [...], ... } }
        student_id = str(data.get('student_id', '')).strip()
        student_name = str(data.get('student_name', '')).strip()
        weeks_data = data.get('weeks', {})

        if not student_id:
            return jsonify({'success': False, 'error': '缺少學員 ID'}), 400

        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)

        try:
            ws = doc.worksheet(SCHEDULE_SHEET_NAME)
        except Exception:
            headers = ['學員姓名', '學員ID'] + [f'W{i}' for i in range(1, 29)]
            ws = doc.add_worksheet(title=SCHEDULE_SHEET_NAME, rows=100, cols=len(headers))
            ws.append_row(headers)

        # 尋找現有列
        all_values = ws.get_all_values()
        target_row = None
        for i, row in enumerate(all_values[1:], start=2):
            if str(row[1] if len(row) > 1 else '').strip() == student_id:
                target_row = i
                break

        # 建立新的一列資料
        new_row = [student_name, student_id]
        for i in range(1, 29):
            stations_list = weeks_data.get(f'W{i}', [])
            # 用逗號串接複數站別
            new_row.append(', '.join(stations_list) if stations_list else '')

        if target_row:
            # 更新現有列
            ws.update(f'A{target_row}', [new_row])
        else:
            # 新增一列
            ws.append_row(new_row)

        logging.info(f"Schedule saved for student {student_id} ({student_name})")
        return jsonify({'success': True, 'msg': f'已儲存 {student_name} 的進度排程'})
    except Exception as e:
        logging.error(f"save_admin_schedule error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/admin/schedule/init_all', methods=['POST'])
def init_all_schedules():
    """對所有學員批量建立預設排程 (W27, W28 = 手動自選，其餘為空)"""
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return jsonify({'success': False, 'error': '僅限管理員'}), 403
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)

        students_records = safe_get_all_records(doc.worksheet('學員名單'))
        students = [(str(r.get('姓名', '')).strip(), str(r.get('學生ID', '')).strip())
                    for r in students_records if str(r.get('姓名', '')).strip()]

        try:
            ws = doc.worksheet(SCHEDULE_SHEET_NAME)
        except Exception:
            headers = ['學員姓名', '學員ID'] + [f'W{i}' for i in range(1, 29)]
            ws = doc.add_worksheet(title=SCHEDULE_SHEET_NAME, rows=100, cols=len(headers))
            ws.append_row(headers)

        all_values = ws.get_all_values()
        existing_ids = {str(row[1] if len(row) > 1 else '').strip() for row in all_values[1:]}

        new_rows = []
        for name, sid in students:
            if sid and sid not in existing_ids:
                row = [name, sid] + ['' for _ in range(26)] + [MANUAL_WEEK_LABEL, MANUAL_WEEK_LABEL]
                new_rows.append(row)

        if new_rows:
            ws.append_rows(new_rows)

        return jsonify({'success': True, 'msg': f'已初始化 {len(new_rows)} 位學員的排程'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/student/schedule', methods=['GET'])
def get_student_schedule():
    """學生查詢自己的本週與全程排程"""
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    try:
        student_info = session.get('student_info', {})
        student_id = str(student_info.get('id', '')).split('.')[0].strip()

        if not student_id:
            return jsonify({'success': False, 'error': '無法辨識學員 ID'}), 400

        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)

        try:
            ws = doc.worksheet(SCHEDULE_SHEET_NAME)
            records = safe_get_all_records(ws)
        except Exception:
            return jsonify({'success': True, 'current_week': None, 'current_stations': [], 'weeks': {}})

        # 找同學的那一筆排程
        student_schedule = None
        for rec in records:
            if str(rec.get('學員ID', '')).strip() == student_id:
                student_schedule = rec
                break

        current_week = get_current_intern_week()
        intern_start = get_internship_start().isoformat()
        current_stations = []
        weeks_map = {}

        if student_schedule:
            for i in range(1, 29):
                val = str(student_schedule.get(f'W{i}', '')).strip()
                stations = [s.strip() for s in val.split(',') if s.strip()] if val else []
                weeks_map[f'W{i}'] = stations

            if current_week and 1 <= current_week <= 28:
                current_stations = weeks_map.get(f'W{current_week}', [])

        return jsonify({
            'success': True,
            'current_week': current_week,
            'current_stations': current_stations,
            'intern_start': intern_start,
            'weeks': weeks_map,
            'manual_weeks': [27, 28]
        })
    except Exception as e:
        logging.error(f"get_student_schedule error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/student/progress_analysis', methods=['GET'])
def get_student_progress_analysis():
    """分析學生 28 週的實習達標狀況 (進度對應)"""
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    try:
        student_info = session.get('student_info', {})
        student_id = str(student_info.get('id', '')).split('.')[0].strip()
        student_name = str(student_info.get('name', '')).strip()

        if not student_id:
            return jsonify({'success': False, 'error': '無法辨識學員 ID'}), 400

        # 1. 取得排程
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        weeks_map = {}
        try:
            ws = doc.worksheet(SCHEDULE_SHEET_NAME)
            records = safe_get_all_records(ws)
            for rec in records:
                if str(rec.get('學員ID', '')).strip() == student_id:
                    for i in range(1, 29):
                        val = str(rec.get(f'W{i}', '')).strip()
                        weeks_map[i] = [s.strip() for s in val.split(',') if s.strip()] if val else []
                    break
        except Exception:
            pass

        # 2. 取得所有成績紀錄 (BigQuery)
        logs = []
        bq_client, project_id = get_bq_client()
        if bq_client:
            from privacy_utils import get_code
            anon_code = get_code(student_name, 'student')
            
            q = f"""
                SELECT station, timestamp
                FROM `{project_id}.grading_data.grading_logs` 
                WHERE (student_id = @code OR student_name = @code) 
                AND (is_deleted = FALSE OR is_deleted IS NULL)
            """
            job_config = bigquery.QueryJobConfig(
                query_parameters=[
                    bigquery.ScalarQueryParameter("code", "STRING", anon_code)
                ]
            )
            res = bq_client.query(q, job_config=job_config).result()
            for r in res:
                # 確保時區處理一致性，BQ 產出通常帶有時區，我們統一轉為 offset-naive 或 local
                t = r.timestamp
                if t and t.tzinfo:
                    t = t.replace(tzinfo=None) + datetime.timedelta(hours=8) # 轉回台北時間基準
                logs.append({'station': str(r.station), 'time': t})

        # 3. 分析 28 週
        analysis = []
        intern_start = get_internship_start()
        current_week = get_current_intern_week()
        
        for i in range(1, 29):
            week_start = intern_start + datetime.timedelta(days=(i-1)*7)
            week_end = week_start + datetime.timedelta(days=6)
            
            # 轉換為 datetime 以進行比較
            ws_dt = datetime.datetime.combine(week_start, datetime.time.min)
            we_dt = datetime.datetime.combine(week_end, datetime.time.max)
            
            assigned_stations = weeks_map.get(i, [])
            
            # 找出落在本週區間內的紀錄
            week_logs = [l for l in logs if l['time'] and ws_dt <= l['time'] <= we_dt]
            
            # 檢查達標狀況
            status = 'pending'
            if i < (current_week or 0):
                if not assigned_stations:
                    status = 'success'
                else:
                    all_hit = True
                    for target in assigned_stations:
                        # 模糊比對站別名稱
                        hit = any(target in l['station'] for l in week_logs)
                        if not hit:
                            all_hit = False
                            break
                    status = 'success' if all_hit else 'fail'
            elif i == current_week:
                if not assigned_stations:
                    status = 'success'
                else:
                    all_hit = True
                    for target in assigned_stations:
                        hit = any(target in l['station'] for l in week_logs)
                        if not hit:
                            all_hit = False
                            break
                    status = 'success' if all_hit else 'pending'
            
            analysis.append({
                'week': i,
                'date_range': f"{week_start.strftime('%m/%d')}~{week_end.strftime('%m/%d')}",
                'stations': assigned_stations,
                'status': status,
                'log_count': len(week_logs)
            })

        return jsonify({
            'success': True,
            'analysis': analysis,
            'current_week': current_week
        })
    except Exception as e:
        logging.error(f"progress_analysis error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/config', methods=['GET'])
def get_config():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
        
    try:
        gc = get_gspread_client()
        if not gc:
            return jsonify({'success': False, 'error': 'Google Sheets backend not configured.'}), 500
            
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        # 取得學員名單 (使用快取)
        def fetch_students():
            try:
                ws_students = doc.worksheet('學員名單')
                return safe_get_all_records(ws_students)
            except: return []
            
        students_records = get_cached_data('students', fetch_students)
        students = []
        for rec in students_records:
            # 支援多種可能的欄位名稱
            name = str(rec.get('姓名') or rec.get('學員姓名') or rec.get('學生姓名') or rec.get('Name') or '').strip()
            sid = str(rec.get('學生ID') or rec.get('學號') or rec.get('ID') or '').split('.')[0].strip()
            
            if name and sid:
                students.append({
                    'id': sid,
                    'name': name,
                    'email': str(rec.get('Email', '')).strip().lower(),
                    'type': str(rec.get('學員類別') or rec.get('職級') or '未分類')
                })
        
        # 取得站別OPA細項
        sheet_stations = doc.worksheet('站別OPA細項')
        stations_records = safe_get_all_records(sheet_stations)
        stations = []
        for rec in stations_records:
            name = str(rec.get('站別', ''))
            if not name:
                continue
                
            body_parts_str = str(rec.get('檢查部位', ''))
            body_parts = [p.strip() for p in body_parts_str.replace('，', ',').split(',')] if body_parts_str else []
            
            station_data = {
                'name': name,
                'body_parts': body_parts,
                'opa1_summary': str(rec.get('OPA1總和評比', '')),
                'opa2_summary': str(rec.get('OPA2總和評比', '')),
                'opa3_summary': str(rec.get('OPA3總和評比', '')),
                'opa1_items': [str(rec.get(f'OPA1_{i}', '')) for i in range(1, 9)],
                'opa2_items': [str(rec.get(f'OPA2_{i}', '')) for i in range(1, 9)],
                'opa3_items': [str(rec.get(f'OPA3_{i}', '')) for i in range(1, 9)],
                'aspect_label': str(rec.get('面向選擇', '')),
                'comment_label': str(rec.get('簡易評語', ''))
            }
            stations.append(station_data)
            
        # 取得信賴等級
        sheet_trust = doc.worksheet('信賴等級描述及轉換')
        trust_records = safe_get_all_records(sheet_trust)
        trust_levels = []
        for rec in trust_records:
            score = str(rec.get('分數', '')).strip()
            level = str(rec.get('信賴等級', '')).strip()
            
            if not score and not level:
                continue
                
            trust_levels.append({
                'score': score,
                'level': level,
                'desc': str(rec.get('描述', '')).strip()
            })
            
        # 4. 取得歷史評分紀錄統計 (優化：改從 BigQuery 抓取以避免 Sheets 429 錯誤)
        student_stats = {}
        try:
            bq_client, project_id = get_bq_client()
            if bq_client:
                from privacy_utils import decode_name
                # 建立 姓名 -> ID 的對應表，因為 BQ 存的是代碼，還原後是姓名，前端需要 ID 作為 Key
                name_to_id = {s['name']: s['id'] for s in students}

                # 抓取 student_id, station, body_part, 以及兩個面向欄位
                q_stats = f"SELECT student_id, station, body_part, aspect1, aspect2 FROM `{project_id}.grading_data.grading_logs` WHERE is_deleted = FALSE OR is_deleted IS NULL"
                stats_results = list(bq_client.query(q_stats).result(timeout=20))
                
                for r in stats_results:
                    # BQ 內的 student_id 現在是匿名代碼
                    anon_id = str(r.student_id).strip()
                    # 還原為真實姓名
                    real_name = decode_name(anon_id)
                    # 轉換為真實 ID (前端 Key)
                    sid = name_to_id.get(real_name)
                    
                    if not sid: continue
                    
                    stn = str(r.station).strip()
                    bpart = str(r.body_part).strip()
                    
                    if not stn or stn == 'None': continue
                    
                    if sid not in student_stats:
                        student_stats[sid] = {'stations': {}}
                    if stn not in student_stats[sid]['stations']:
                        student_stats[sid]['stations'][stn] = {
                            'count': 0,
                            'body_parts': {},
                            'aspects': {}
                        }
                    
                    stn_data = student_stats[sid]['stations'][stn]
                    stn_data['count'] += 1
                    
                    if bpart and bpart != 'None':
                        stn_data['body_parts'][bpart] = stn_data['body_parts'].get(bpart, 0) + 1
                    
                    # 統計面向 (在 BQ 中 aspect1, aspect2 存的是面向編號 1~5)
                    for asp_val in [r.aspect1, r.aspect2]:
                        if asp_val:
                            # 提取數字部分
                            import re
                            m = re.search(r'\d+', str(asp_val))
                            if m:
                                asp_num = m.group(0)
                                stn_data['aspects'][asp_num] = stn_data['aspects'].get(asp_num, 0) + 1
        except Exception as bq_err:
            logging.error(f"Error calculating stats from BQ: {bq_err}")
            
        return jsonify({
            'success': True,
            'students': students,
            'stations': stations,
            'trust_levels': trust_levels,
            'student_stats': student_stats
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/student_stats', methods=['GET'])
def get_student_stats():
    # [Phase 6] 支援分享模式存取
    is_shared = session.get('is_shared_view', False)
    if is_shared:
        target_id = session.get('shared_student_id')
        target_name = session.get('shared_student_name')
    else:
        user = session.get('user')
        if not user:
            return jsonify({'success': False, 'error': 'Unauthorized'}), 401
        if session.get('current_role') != 'student':
            return jsonify({'success': False, 'error': 'Forbidden'}), 403
        student_info = session.get('student_info')
        if not student_info:
            return jsonify({'success': False, 'error': 'No info'}), 400
        target_id = str(student_info.get('id', '')).split('.')[0].strip()
        target_name = str(student_info.get('name', '')).strip()

    try:
        from privacy_utils import get_code, decode_name
        anon_s_code = get_code(target_name, 'student')
        
        records_by_station = {}
        bq_client, project_id = get_bq_client()
        if bq_client:
            # BQ 內 student_id 與 student_name 現在存的都是匿名代碼
            q = f"SELECT station, timestamp, opa1_sum, opa2_sum, opa3_sum, teacher_name, aspect1, aspect2, comment FROM `{project_id}.grading_data.grading_logs` WHERE (student_id = @code OR student_name = @code) AND (is_deleted = FALSE OR is_deleted IS NULL) ORDER BY timestamp ASC"
            job_config = bigquery.QueryJobConfig(query_parameters=[
                bigquery.ScalarQueryParameter("code", "STRING", anon_s_code)
            ])
            res = bq_client.query(q, job_config=job_config).result(timeout=20)
            def to_int(v):
                try: return int(v)
                except: return None
            for r in res:
                stn = str(r.station or 'Unknown').strip()
                if stn not in records_by_station: records_by_station[stn] = []
                # 將 BQ 內的教師代碼還原為姓名顯示
                teacher_display = decode_name(str(r.teacher_name or ''))
                
                records_by_station[stn].append({
                    'time': r.timestamp.strftime('%Y/%m/%d %H:%M:%S') if r.timestamp else '',
                    'teacher': teacher_display or '未知',
                    'opa1': to_int(r.opa1_sum), 'opa2': to_int(r.opa2_sum), 'opa3': to_int(r.opa3_sum),
                    'aspect1': r.aspect1 or '', 'aspect2': r.aspect2 or '', 'comment': r.comment or ''
                })
            return jsonify({'success': True, 'data': records_by_station})
        return jsonify({'success': False, 'error': 'BQ client missing'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/submit_grade', methods=['POST'])
def submit_grade():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    data = request.json
    try:
        import datetime
        timestamp = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
        teacher_email = user.get('email', '')
        teacher_name = teacher_email
        
        # 取得教師姓名 (優先使用快取)
        def fetch_teachers():
            try:
                gc = get_gspread_client()
                sheet_id = os.getenv("GOOGLE_SHEET_ID")
                doc = gc.open_by_key(sheet_id)
                sheet_teachers = doc.worksheet('教師名單')
                return safe_get_all_records(sheet_teachers)
            except: return []
            
        teachers_data = get_cached_data('teachers', fetch_teachers)
        for rec in teachers_data:
            if str(rec.get('教師_Email', '')).strip().lower() == teacher_email.lower():
                teacher_name = str(rec.get('教師姓名', '')).strip() or teacher_name
                break

        student_id = data.get('student_id', '')
        student_name = data.get('student_name', '')
        station = data.get('station', '')
        body_part = data.get('body_part', '')
        opa1_sum = data.get('opa1_sum', '')
        opa2_sum = data.get('opa2_sum', '')
        opa3_sum = data.get('opa3_sum', '')
        opa1_items = (data.get('opa1_items', [''] * 8) + [''] * 8)[:8]
        opa2_items = (data.get('opa2_items', [''] * 8) + [''] * 8)[:8]
        opa3_items = (data.get('opa3_items', [''] * 8) + [''] * 8)[:8]
        aspect1 = data.get('aspect1', '')
        aspect2 = data.get('aspect2', '')
        comment = data.get('comment', '')

        # --- 1. 優先寫入 BigQuery (穩定性高，不限流) ---
        bq_success = False
        try:
            bq_client, project_id = get_bq_client()
            if bq_client:
                dt_iso = datetime.datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S").isoformat()
                grade_insert = [{
                    'student_id': student_id, 'student_name': student_name, 'station': station,
                    'body_part': body_part, 'timestamp': dt_iso, 'teacher_name': teacher_name,
                    'opa1_sum': opa1_sum, 'opa2_sum': opa2_sum, 'opa3_sum': opa3_sum,
                    'opa1_items': ",".join(map(str, opa1_items)),
                    'opa2_items': ",".join(map(str, opa2_items)),
                    'opa3_items': ",".join(map(str, opa3_items)),
                    'aspect1': aspect1, 'aspect2': aspect2, 'comment': comment,
                    'is_deleted': False
                }]
                errors = bq_client.insert_rows_json(f"{project_id}.grading_data.grading_logs", grade_insert)
                if not errors: bq_success = True
                else: logging.error(f"BQ Insert Errors: {errors}")
        except Exception as e:
            logging.error(f"BQ Write FATAL: {e}")

        # --- 2. 寫入 Google Sheets (作為鏡像備份，允許低機率失敗) ---
        try:
            gc = get_gspread_client()
            sheet_id = os.getenv("GOOGLE_SHEET_ID")
            doc = gc.open_by_key(sheet_id)
            sheet = doc.worksheet('評分記錄')
            row = [student_id, student_name, station, body_part, timestamp, teacher_name, opa1_sum, opa2_sum, opa3_sum] + opa1_items + opa2_items + opa3_items + [aspect1, aspect2, comment]
            sheet.append_row(row, table_range="A1")
        except Exception as e:
            logging.warning(f"Sheets Write Failed (but BQ might be OK): {e}")
            if not bq_success:
                # 如果連 BQ 都沒成功，才報錯給教師
                return jsonify({'success': False, 'error': '資料庫寫入失敗，請稍後再試。'}), 500

        return jsonify({'success': True, 'msg': '評分已儲存' if bq_success else '評分已記錄至備用路徑'})
    except Exception as e:
        logging.error(f"Submit Grade Error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

def aggregate_student_report_data(doc, bq_client, project_id):
    import time
    ws_students = doc.worksheet('學員名單')
    student_records = safe_get_all_records(ws_students)
    time.sleep(1) 
    hours_map = {}
    try:
        ws_attendance = doc.worksheet('上下班打卡記錄')
        attendance_records = safe_get_all_records(ws_attendance)
        time.sleep(1)
        for r in attendance_records:
            name = str(r.get('學生', '')).strip()
            if not name: continue
            try:
                t_in = datetime.datetime.strptime(r.get('簽到時間', ''), "%Y-%m-%d %H:%M:%S")
                t_out = datetime.datetime.strptime(r.get('簽退時間', ''), "%Y-%m-%d %H:%M:%S")
                hours_map[name] = hours_map.get(name, 0) + (t_out - t_in).total_seconds() / 3600.0
            except: continue
    except: pass
    opa_stats = {}
    if bq_client:
        q = f"SELECT student_name, station, AVG(opa1_sum + opa2_sum + opa3_sum) as total_avg, COUNT(*) as count FROM `{project_id}.grading_data.grading_logs` WHERE is_deleted = FALSE OR is_deleted IS NULL GROUP BY student_name, station"
        res = bq_client.query(q).result()
        for r in res:
            name = str(r.student_name).strip()
            if name not in opa_stats: opa_stats[name] = {}
            opa_stats[name][r.station] = {'avg': round(r.total_avg or 0, 2), 'count': r.count}
    
    def clean_ceep_name(raw_name):
        import re
        m = re.search(r'^(.+?)\s*[（\(]', str(raw_name))
        return m.group(1).strip() if m else str(raw_name).strip()
    
    # --- 3. DOPS 數據 (從 BigQuery 讀取) ---
    dops_stats = {}
    if bq_client:
        try:
            q_dops = f"SELECT student_name, station, score, feedback FROM `{project_id}.grading_data.dops_logs`"
            res_dops = bq_client.query(q_dops).result()
            for r in res_dops:
                name = clean_ceep_name(r.student_name)
                stn = str(r.station).strip()
                if name not in dops_stats: dops_stats[name] = {}
                if stn not in dops_stats[name]: dops_stats[name][stn] = {'scores': [], 'feedbacks': []}
                dops_stats[name][stn]['scores'].append(r.score)
                if r.feedback: dops_stats[name][stn]['feedbacks'].append(r.feedback)
        except Exception as e:
            logging.error(f"BQ DOPS Read Error: {e}")

    # --- 4. Mini-CEX 數據 (從 BigQuery 讀取) ---
    mcex_stats = {}
    if bq_client:
        try:
            q_mcex = f"SELECT student_name, station, score, feedback FROM `{project_id}.grading_data.minicex_logs`"
            res_mcex = bq_client.query(q_mcex).result()
            for r in res_mcex:
                name = clean_ceep_name(r.student_name)
                # MiniCEX 站別通常包含 "-", 需要清理
                import re
                station_match = re.search(r'^(.+?)-', str(r.station))
                stn = station_match.group(1).strip() if station_match else str(r.station).strip()
                
                if name not in mcex_stats: mcex_stats[name] = {}
                if stn not in mcex_stats[name]: mcex_stats[name][stn] = {'scores': [], 'feedbacks': []}
                mcex_stats[name][stn]['scores'].append(r.score)
                if r.feedback: mcex_stats[name][stn]['feedbacks'].append(r.feedback)
        except Exception as e:
            logging.error(f"BQ MiniCEX Read Error: {e}")

    # 5. 進度達標率判定邏輯
    progress_map = {}
    try:
        ws_sched = doc.worksheet(SCHEDULE_SHEET_NAME)
        sched_records = safe_get_all_records(ws_sched)
        intern_start = get_internship_start()
        current_week = get_current_intern_week() or 28
        
        q_all = f"SELECT student_name, station, timestamp FROM `{project_id}.grading_data.grading_logs` WHERE is_deleted = FALSE OR is_deleted IS NULL"
        all_logs = list(bq_client.query(q_all).result())
        
        for rec in sched_records:
            s_name = str(rec.get('學員姓名', '')).strip()
            if not s_name: continue
            student_logs = [l for l in all_logs if str(l.student_name).strip() == s_name]
            success_weeks = 0
            total_relevant_weeks = min(current_week, 28)
            for i in range(1, total_relevant_weeks + 1):
                w_start = intern_start + datetime.timedelta(days=(i-1)*7)
                w_end = w_start + datetime.timedelta(days=6)
                ws_dt = datetime.datetime.combine(w_start, datetime.time.min)
                we_dt = datetime.datetime.combine(w_end, datetime.time.max)
                val = str(rec.get(f'W{i}', '')).strip()
                assigned = [s.strip() for s in val.split(',') if s.strip()]
                if not assigned:
                    success_weeks += 1; continue
                week_logs = [l for l in student_logs if ws_dt <= l.timestamp.replace(tzinfo=None) + datetime.timedelta(hours=8) <= we_dt]
                all_hit = True
                for target in assigned:
                    if not any(target in str(l.station) for l in week_logs):
                        all_hit = False; break
                if all_hit: success_weeks += 1
            rate = (success_weeks/total_relevant_weeks)*100 if total_relevant_weeks > 0 else 0
            progress_map[s_name] = f"{round(rate, 1)}%"
    except: pass

    # 6. 彙整結果 (Wide Format)
    all_stations = set()
    for s_map in opa_stats.values(): all_stations.update(s_map.keys())
    for s_map in dops_stats.values(): all_stations.update(s_map.keys())
    for s_map in mcex_stats.values(): all_stations.update(s_map.keys())
    sorted_stations = sorted(list(all_stations))

    rows = []
    for s in student_records:
        name = str(s.get('姓名', '')).strip()
        if not name: continue
        
        row = {
            '學員ID': s.get('學生ID', ''),
            '姓名': name,
            '學員類別': s.get('學員類別', ''),
            '實習總時數': round(hours_map.get(name, 0), 1),
            '進度達標率': progress_map.get(name, '0%')
        }
        
        for stn in sorted_stations:
            # OPA 數據 (原有邏輯)
            opa = opa_stats.get(name, {}).get(stn, {'avg': '-', 'count': 0})
            row[f'[{stn}] OPA平均'] = opa['avg']
            row[f'[{stn}] OPA筆數'] = opa['count']
            
            # DOPS 數據 (更新後)
            d_data = dops_stats.get(name, {}).get(stn, {'scores': [], 'feedbacks': []})
            d_scores = d_data['scores']
            row[f'[{stn}] DOPS平均'] = round(sum(d_scores)/len(d_scores), 2) if d_scores else '-'
            row[f'[{stn}] DOPS回饋'] = " | ".join(d_data['feedbacks']) if d_data['feedbacks'] else ''
            
            # Mini-CEX 數據 (更新後)
            m_data = mcex_stats.get(name, {}).get(stn, {'scores': [], 'feedbacks': []})
            m_scores = m_data['scores']
            row[f'[{stn}] MiniCEX平均'] = round(sum(m_scores)/len(m_scores), 2) if m_scores else '-'
            row[f'[{stn}] MiniCEX回饋'] = " | ".join(m_data['feedbacks']) if m_data['feedbacks'] else ''
            
        rows.append(row)
    return rows

@app.route('/api/admin/export_excel', methods=['GET'])
def admin_export_excel():
    """管理員匯出實習成績與進度總表 (Excel)"""
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        bq_client, project_id = get_bq_client()
        
        # 呼叫統計函式
        final_data = aggregate_student_report_data(doc, bq_client, project_id)
            
        df = pd.DataFrame(final_data)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='實習成績總表')
        output.seek(0)
        
        filename = f"Internship_Report_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        logging.error(f"Export Excel Error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/student_gamification', methods=['GET'])
def get_student_gamification():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    if session.get('current_role') != 'student':
        return jsonify({'success': False, 'error': 'Forbidden'}), 403
        
    student_info = session.get('student_info')
    if not student_info:
        return jsonify({'success': False, 'error': 'No student info found.'}), 400
        
    try:
        from gamification import get_student_gamification_data, parse_scoring_config
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        # 準備快取數據
        def f_students(): return doc.worksheet('學員名單').get_all_values()
        def f_epa(): return doc.worksheet('各類別EPA需求').get_all_values()
        def f_rooms(): return doc.worksheet('檢查室清單').get_all_values()
        def f_score(): return parse_scoring_config(doc)
        
        cached = {
            'students': get_cached_data('students', f_students),
            'epa_requirements': get_cached_data('epa_requirements', f_epa),
            'exam_rooms': get_cached_data('exam_rooms', f_rooms),
            'scoring_config': get_cached_data('scoring_config', f_score)
        }
        
        data = get_student_gamification_data(gc, doc, student_info, cached_data=cached)
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/leaderboard', methods=['GET'])
def get_leaderboard():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
        
    try:
        from gamification import get_leaderboard_data, parse_scoring_config
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        # 準備快取數據
        def f_students(): return doc.worksheet('學員名單').get_all_values()
        def f_epa(): return doc.worksheet('各類別EPA需求').get_all_values()
        def f_rooms(): return doc.worksheet('檢查室清單').get_all_values()
        def f_score(): return parse_scoring_config(doc)
        
        cached = {
            'students': get_cached_data('students', f_students),
            'epa_requirements': get_cached_data('epa_requirements', f_epa),
            'exam_rooms': get_cached_data('exam_rooms', f_rooms),
            'scoring_config': get_cached_data('scoring_config', f_score)
        }
        
        data = get_leaderboard_data(gc, doc, cached_data=cached)
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/student_attendance', methods=['GET'])
def get_student_attendance():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
        
    student_info = session.get('student_info')
    if not student_info:
        return jsonify({'success': False, 'error': 'No student info found.'}), 400
        
    try:
        from privacy_utils import get_code, decode_name
        # 將真實姓名轉換為去識別化代碼
        anon_s_code = get_code(student_info['name'], 'student')
        
        history = []
        bq_client, project_id = get_bq_client()
        if bq_client:
            # BQ 內的 student_name 現在存的是代碼
            q = f"SELECT sub_room, check_in_time, check_out_time, teacher_name, co_teacher FROM `{project_id}.grading_data.attendance_daily_summary` WHERE student_name = @code ORDER BY check_in_time DESC"
            job_config = bigquery.QueryJobConfig(
                query_parameters=[bigquery.ScalarQueryParameter("code", "STRING", anon_s_code)]
            )
            # Add explicit timeout of 20 seconds to prevent hanging the Flask thread
            query_job = bq_client.query(q, job_config=job_config)
            res = query_job.result(timeout=20)
            
            for r in res:
                # Convert row to dict for safer access across different versions of BQ client
                row_dict = dict(r)
                # 將 BQ 內的教師代碼還原為姓名顯示
                teacher_display = decode_name(row_dict.get('teacher_name', ''))
                
                history.append({
                    'room': row_dict.get('sub_room', ''),
                    'check_in': row_dict.get('check_in_time').strftime('%Y/%m/%d %H:%M:%S') if row_dict.get('check_in_time') else '',
                    'check_out': row_dict.get('check_out_time').strftime('%Y/%m/%d %H:%M:%S') if row_dict.get('check_out_time') else '',
                    'teacher': teacher_display,
                    'co_teacher': row_dict.get('co_teacher', ''),
                    'is_complete': bool(row_dict.get('check_in_time') and row_dict.get('check_out_time'))
                })
        
        return jsonify({'success': True, 'data': history})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# [Phase 6] 產生分享連結 API
@app.route('/api/admin/generate_share_link/<student_id>', methods=['POST'])
def generate_share_link(student_id):
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    try:
        from credentials_utils import get_gspread_client
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        ws_students = doc.worksheet('學員名單')
        students = safe_get_all_records(ws_students)
        student_name = "Unknown"
        for s in students:
            if str(s.get('學生ID', '')).split('.')[0] == str(student_id):
                student_name = s.get('姓名', 'Unknown')
                break
        
        try:
            ws = doc.worksheet('分享連結管理')
        except:
            ws = doc.add_worksheet(title='分享連結管理', rows=100, cols=10)
            ws.append_row(['學員ID', '學員姓名', '權杖', '啟用狀態', '建立時間'])
            
        token = str(uuid.uuid4())[:16]
        all_vals = ws.get_all_values()
        found_row = -1
        for i, row in enumerate(all_vals):
            if i == 0: continue
            if len(row) > 0 and str(row[0]).split('.')[0] == str(student_id):
                found_row = i + 1
                break
        
        timestamp = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
        if found_row != -1:
            ws.update_cell(found_row, 3, token)
            ws.update_cell(found_row, 4, 'TRUE')
            ws.update_cell(found_row, 5, timestamp)
        else:
            ws.append_row([student_id, student_name, token, 'TRUE', timestamp])
            
        share_url = f"{request.host_url.rstrip('/')}/share/{token}"
        return jsonify({'success': True, 'token': token, 'share_url': share_url})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# [Phase 4] 視覺化報表數據 API (Pro Dashboard 使用)
@app.route('/api/student_report_data')
def get_student_report_data():
    is_shared = session.get('is_shared_view', False)
    is_preview = session.get('is_preview_mode', False)
    
    if is_shared:
        target_id = session.get('shared_student_id')
        target_name = session.get('shared_student_name')
    elif is_preview:
        preview_info = session.get('preview_student_info', {})
        target_id = str(preview_info.get('id', '')).split('.')[0].strip()
        target_name = str(preview_info.get('name', '')).strip()
    else:
        user = session.get('user')
        if not user: return jsonify({'success': False, 'error': 'Unauthorized'}), 401
        student_info = session.get('student_info')
        if not student_info: return jsonify({'success': False, 'error': 'No info'}), 400
        target_id = str(student_info.get('id', '')).split('.')[0].strip()
        target_name = str(student_info.get('name', '')).strip()

    try:
        from privacy_utils import get_code
        # 將真實身分轉換為去識別化代碼以利 BQ 查詢
        anon_code = get_code(target_name, 'student')
        
        bq_client, project_id = get_bq_client()
        if not bq_client: return jsonify({'success': False, 'error': 'BQ Error'}), 500
        
        # 1. 抓取 EPA Grading Logs
        q_epa = f"SELECT teacher_name, station, timestamp, opa1_sum, opa2_sum, opa3_sum, aspect1, aspect2, comment FROM `{project_id}.grading_data.grading_logs` WHERE (student_id = @code OR student_name = @code) AND (is_deleted = FALSE OR is_deleted IS NULL)"
        job_config = bigquery.QueryJobConfig(query_parameters=[bigquery.ScalarQueryParameter("code", "STRING", anon_code)])
        epa_results = bq_client.query(q_epa, job_config=job_config).result(timeout=20)
        
        raw_logs = []
        for r in epa_results:
            teacher_display = decode_name(str(r.teacher_name or ''))
            raw_logs.append({
                'type': 'EPA',
                'teacher': teacher_display or '未知',
                'station': str(r.station or '未知'),
                'date': r.timestamp.strftime('%Y-%m-%d'),
                'timestamp': r.timestamp.isoformat(),
                'm_d': r.timestamp.strftime('%m/%d'),
                'opa1': float(r.opa1_sum or 0) if str(r.opa1_sum).replace('.','').isdigit() else 0,
                'opa2': float(r.opa2_sum or 0) if str(r.opa2_sum).replace('.','').isdigit() else 0,
                'opa3': float(r.opa3_sum or 0) if str(r.opa3_sum).replace('.','').isdigit() else 0,
                'aspect1': str(r.aspect1 or '未評定').strip(),
                'aspect2': float(r.aspect2 or 0) if str(r.aspect2).replace('.','').isdigit() else 0,
                'comment': str(r.comment or '無質性回饋').strip()
            })

        # 2. 抓取 CEEP DOPS
        q_dops = f"SELECT timestamp, station, score, feedback FROM `{project_id}.grading_data.dops_logs` WHERE student_name = @code"
        dops_results = bq_client.query(q_dops, job_config=job_config).result(timeout=20)
        for r in dops_results:
            raw_logs.append({
                'type': 'DOPS',
                'teacher': 'CEEP 系統',
                'station': str(r.station or '未知'),
                'date': r.timestamp.strftime('%Y-%m-%d') if r.timestamp else '未知',
                'timestamp': r.timestamp.isoformat() if r.timestamp else '',
                'm_d': r.timestamp.strftime('%m/%d') if r.timestamp else '',
                'score': float(r.score or 0),
                'comment': str(r.feedback or '').strip()
            })

        # 3. 抓取 CEEP MiniCEX
        q_cex = f"SELECT timestamp, station, score, feedback FROM `{project_id}.grading_data.minicex_logs` WHERE student_name = @code"
        cex_results = bq_client.query(q_cex, job_config=job_config).result(timeout=20)
        for r in cex_results:
            raw_logs.append({
                'type': 'MiniCEX',
                'teacher': 'CEEP 系統',
                'station': str(r.station or '未知'),
                'date': r.timestamp.strftime('%Y-%m-%d') if r.timestamp else '未知',
                'timestamp': r.timestamp.isoformat() if r.timestamp else '',
                'm_d': r.timestamp.strftime('%m/%d') if r.timestamp else '',
                'score': float(r.score or 0),
                'comment': str(r.feedback or '').strip()
            })

        # 4. 抓取 CEEP 教學記錄
        q_teach = f"SELECT timestamp, station, content, feedback FROM `{project_id}.grading_data.teaching_logs` WHERE student_name = @code"
        teach_results = bq_client.query(q_teach, job_config=job_config).result(timeout=20)
        for r in teach_results:
            raw_logs.append({
                'type': '教學記錄',
                'teacher': 'CEEP 系統',
                'station': str(r.station or '未知'),
                'date': r.timestamp.strftime('%Y-%m-%d') if r.timestamp else '未知',
                'timestamp': r.timestamp.isoformat() if r.timestamp else '',
                'm_d': r.timestamp.strftime('%m/%d') if r.timestamp else '',
                'content': str(r.content or '').strip(),
                'comment': str(r.feedback or '').strip()
            })

        # 統一按時間降序排列 (最新在前)
        raw_logs.sort(key=lambda x: x.get('timestamp', ''), reverse=True)

        return jsonify({
            'success': True,
            'raw_logs': raw_logs
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# [Phase 6] 分享連結驗證輔助函數
def get_share_info_by_token(token):
    try:
        from credentials_utils import get_gspread_client
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        try:
            ws = doc.worksheet('分享連結管理')
        except:
            ws = doc.add_worksheet(title='分享連結管理', rows=100, cols=10)
            ws.append_row(['學員ID', '學員姓名', '權杖', '啟用狀態', '建立時間'])
            
        all_data = safe_get_all_records(ws)
        for row in all_data:
            if str(row.get('權杖', '')).strip() == str(token).strip():
                if str(row.get('啟用狀態', '')).strip().upper() == 'TRUE':
                    return {
                        'id': str(row.get('學員ID', '')),
                        'name': str(row.get('學員姓名', ''))
                    }
        return None
    except Exception as e:
        print(f"Share token validation error: {e}")
        return None

@app.route('/student/pro_report')
def student_pro_report():
    is_shared = session.get('is_shared_view', False)
    is_preview = session.get('is_preview_mode', False)
    
    # 優先從預覽或分享資訊中獲取
    student_info = session.get('preview_student_info') or session.get('student_info')
    
    if not is_shared and not is_preview and not session.get('user'):
        return redirect(url_for('login'))
    return render_template('student_report_v2.html', student_info=student_info)

@app.route('/share/<token>')
def view_shared_dashboard(token):
    share_info = get_share_info_by_token(token)
    if not share_info:
        return f"<h1>分享連結無效或已過期</h1><p>請聯繫管理員重新產生連結。</p>", 403
    
    # 設定分享視圖參數，但不強制覆蓋現有 user session (如果是管理員在預覽)
    session['is_shared_view'] = True
    # 補抓學員類別以顯示基準線
    try:
        gc = get_gspread_client()
        doc = gc.open_by_key(os.getenv("GOOGLE_SHEET_ID"))
        sheet = doc.worksheet('學員名單')
        records = safe_get_all_records(sheet)
        target = next((r for r in records if str(r.get('學生ID', '')).split('.')[0] == str(share_info['id'])), None)
        if target:
            share_info['type'] = str(target.get('學員類別', ''))
    except: pass

    session['student_info'] = share_info
    
    # 只有在未登入狀態下才賦予虛擬查核員身份
    if not session.get('user'):
        session['user'] = {'name': f"外部查核員 (觀看 {share_info['name']})", 'email': 'guest@shared.view', 'picture': ''}
        session['current_role'] = 'student'
    
    return redirect(url_for('student_pro_report'))

@app.route('/admin/view_report/<student_id>')
def admin_view_report(student_id):
    if not session.get('is_admin'):
        return redirect(url_for('login'))
    
    # 獲取學員列表以讀取正確的姓名
    try:
        from credentials_utils import get_gspread_client
        gc = get_gspread_client()
        if not gc: raise Exception("無法獲取 Google Sheets 客戶端")
        
        doc = gc.open_by_key(os.getenv("GOOGLE_SHEET_ID"))
        # 靈活匹配工作表名稱，避免編碼或空格問題
        all_ws = doc.worksheets()
        sheet = next((w for w in all_ws if "學員名單" in w.title), None)
        
        if not sheet:
            titles = [w.title for w in all_ws]
            return f"找不到工作表 '學員名單'。現有的表包含: {titles}", 404
            
        records = safe_get_all_records(sheet)
        # ID 匹配：同樣採用靈活匹配
        target = next((r for r in records if str(r.get('學生ID', '')).split('.')[0] == str(student_id).split('.')[0]), None)
        
        if target:
            session['preview_student_info'] = {
                'id': str(target.get('學生ID', '')),
                'name': str(target.get('姓名', '')),
                'type': str(target.get('學員類別', ''))
            }
            session['is_preview_mode'] = True
            return redirect(url_for('student_pro_report'))
    except Exception as e:
        return f"載入失敗: {e}", 500
    
    return "找不到該學員資訊", 404

@app.route('/admin')
def admin_portal():
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return render_template('error.html', message='權限不足：此頁面僅限系統管理員進入。'), 403
    
    # 確保跳回管理中心時，清除預覽模式與學員身分切換的殘留
    session.pop('is_preview_mode', None)
    session.pop('preview_student_info', None)
    session['current_role'] = 'teacher' # 強制切換回老師/管理者視角
    
    # 傳遞 Sheets ID 供前端產生連結
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    sheet_feedback = "112l_e3WKbIkFYj58nv8LRTYEvfyDpXMh-NcSe98T07w"
    
    # 獲取學員名單 (使用快取)
    students = []
    try:
        gc = get_gspread_client()
        doc = gc.open_by_key(sheet_id)
        
        def fetch_students_admin():
            try:
                # 嘗試多種可能的分頁名稱
                ws = None
                for name in ['學生名單', '學員名單', '學員名冊']:
                    try:
                        ws = doc.worksheet(name)
                        break
                    except: continue
                if not ws: return []
                return safe_get_all_records(ws)
            except: return []
            
        records = get_cached_data('students', fetch_students_admin)
        for r in records:
            # 支援多種可能的欄位名稱
            s_name = str(r.get('姓名') or r.get('學員姓名') or r.get('學生姓名') or r.get('Name') or '').strip()
            s_id = str(r.get('學生ID') or r.get('學號') or r.get('ID') or '').split('.')[0].strip()
            
            if s_name and s_id:
                students.append({'name': s_name, 'id': s_id})
        
        logging.info(f"Admin Portal: Loaded {len(students)} students for dropdowns.")
    except Exception as e:
        logging.error(f"Admin fetch students error: {e}")

    return render_template('admin_portal.html', 
                           user=user, 
                           roles=session.get('roles', []),
                           sheet_main=sheet_id,
                           sheet_feedback=sheet_feedback,
                           students=students)

@app.route('/admin/schedule')
def admin_schedule():
    """學生進度排程管理頁面"""
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return render_template('error.html', message='權限不足：此頁面僅限系統管理員進入。'), 403
    return render_template('schedule_manager.html', user=user)


@app.route('/admin/simulate_student', methods=['POST'])
def simulate_student():
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return jsonify({'success': False, 'error': 'Unauthorized'}), 403
    
    student_id = request.form.get('student_id')
    if not student_id:
        return redirect('/admin')
        
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        sheet_students = doc.worksheet('學員名單')
        records = safe_get_all_records(sheet_students)
        
        target_student = None
        for r in records:
            if str(r.get('學生ID', '')).strip() == student_id:
                target_student = {
                    'id': str(r.get('學生ID', '')).strip(),
                    'name': str(r.get('姓名', '')),
                    'type': str(r.get('學員類別', '')),
                    'gender': str(r.get('性別', ''))
                }
                break
        
        if target_student:
            session['student_info'] = target_student
            session['current_role'] = 'student'
            session['is_simulating'] = True
            return redirect('/')
        else:
            return "找不到該學員資料", 404
            
    except Exception as e:
        return f"模擬失敗: {str(e)}", 500

@app.route('/admin/stop_simulation')
def stop_simulation():
    if session.get('is_simulating'):
        session['is_simulating'] = False
        session['current_role'] = 'teacher'
        # 清除模擬的學員資訊，重新觸發權限檢查（或者保留原本的，但 current_role 已改）
        # 為了安全，重新導向回 admin 會觸發一般的 session 檢查
    return redirect('/admin')

@app.route('/admin/attendance_monitor')
def admin_attendance_monitor():
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return render_template('error.html', message='權限不足：此頁面僅限系統管理員進入。'), 403
    return render_template('admin_attendance_monitor.html', user=user, roles=session.get('roles', []))

@app.route('/admin/course_qrcodes')
def admin_course_qrcodes():
    user = session.get('user')
    # 改為嚴格管理員檢查
    if not user or not session.get('is_admin'):
        return render_template('error.html', message='權限不足：課程條碼產生器僅限管理員存取。'), 403
    
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        ws = doc.worksheet('早8課程簽到(含放腫、全人教學)')
        courses = safe_get_all_records(ws)
        return render_template('course_qrcodes.html', courses=courses, roles=session.get('roles', []))
    except Exception as e:
        return f"Error loading courses: {e}", 500

@app.route('/course_checkin')
def course_checkin_landing():
    # 學員掃描 QR 後進入此頁面進行登入/確認
    user = session.get('user')
    if not user:
        session['next_url'] = request.url
        return redirect('/login')
    
    course_name = request.args.get('course')
    return render_template('course_checkin_confirm.html', course_name=course_name, user=user)

@app.route('/api/course_checkin', methods=['POST'])
def api_course_checkin():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
        
    student_info = session.get('student_info', {})
    data = request.json
    course_name = data.get('course_name')
    is_manual = data.get('is_manual', False)
    
    if not course_name:
        return jsonify({'success': False, 'error': 'Missing course name'}), 400

    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        ws = doc.worksheet('早8課程簽到(含放腫、全人教學)')
        course_data = safe_get_all_records(ws)
        
        target_course = None
        for c in course_data:
            if c.get('課程名稱') == course_name:
                target_course = c
                break
        
        if not target_course:
            return jsonify({'success': False, 'error': '找不到該課程資訊'}), 404

        # 1. 時間驗證 (管理員補登或特定模式可跳過)
        if not is_manual:
            try:
                import datetime
                # 假設試算表格式: 上課日期 (YYYY/MM/DD), 開始時間 (HH:MM)
                date_str = str(target_course.get('上課日期', '')).strip()
                time_str = str(target_course.get('開始時間', '')).strip()
                # 靈活處理分隔符號
                date_str = date_str.replace('-', '/')
                start_dt = datetime.datetime.strptime(f"{date_str} {time_str}", "%Y/%m/%d %H:%M")
                
                # 台灣時間 (UTC+8)
                now = datetime.datetime.utcnow() + datetime.timedelta(hours=8)
                
                # 限制：前 15 分 ~ 後 5 分
                early_boundary = start_dt - datetime.timedelta(minutes=15)
                late_boundary = start_dt + datetime.timedelta(minutes=5)
                
                if now < early_boundary:
                    return jsonify({'success': False, 'error': f'簽到尚未開始。請於 {early_boundary.strftime("%H:%M")} 後再試。'}), 403
                if now > late_boundary:
                    return jsonify({'success': False, 'error': '簽到已逾時結束。請洽管理員補登。'}), 403
            except Exception as e:
                import logging
                logging.error(f"Time validation error for course {course_name}: {e}")
                return jsonify({'success': False, 'error': '系統無法驗證課程時間 (需檢查日期/時間格式)，請聯繫老師手動補登。'}), 500

        # 2. 檢查重複簽到
        bq_client, project_id = get_bq_client()
        student_id = str(student_info.get('id', '')).split('.')[0].strip()
        if not student_id:
            return jsonify({'success': False, 'error': '無法辨識學員 ID，請重新登入。'}), 400
        
        q_dup = f"SELECT count(*) as cnt FROM `{project_id}.grading_data.course_checkins` WHERE CAST(student_id AS STRING) = @sid AND course_name = @cname"
        job_config = bigquery.QueryJobConfig(
            query_parameters=[
                bigquery.ScalarQueryParameter("sid", "STRING", student_id),
                bigquery.ScalarQueryParameter("cname", "STRING", course_name)
            ]
        )
        res_dup = bq_client.query(q_dup, job_config=job_config).result()
        for r in res_dup:
            if r.cnt > 0:
                return jsonify({'success': False, 'error': '您已完成此課程簽到，請勿重複掃描。'}), 400

        # 3. 執行簽到
        try:
            hours = float(target_course.get('時數', 0))
        except:
            hours = 0.0
            
        now_iso = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).isoformat()
        
        row_c = [{
            'student_id': student_id,
            'student_name': student_info.get('name', 'Unknown'),
            'course_name': course_name,
            'hours': hours,
            'is_manual': is_manual,
            'timestamp': now_iso
        }]
        errors = bq_client.insert_rows_json(f"{project_id}.grading_data.course_checkins", row_c)
        if errors:
            return jsonify({'success': False, 'error': f'BQ 寫入失敗: {str(errors)}'}), 500
            
        return jsonify({'success': True, 'msg': f'簽到成功！獲得 {hours} 小時時數。'})
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/admin/manual_checkin', methods=['POST'])
def api_admin_manual_checkin():
    user = session.get('user')
    if not user or 'teacher' not in session.get('roles', []):
        return jsonify({'success': False, 'error': 'Forbidden'}), 403
    
    data = request.json
    target_sid = str(data.get('student_id', '')).split('.')[0].strip()
    target_name = data.get('student_name', '')
    course_name = data.get('course_name', '')
    hours = float(data.get('hours', 0))
    
    bq_client, project_id = get_bq_client()
    now_iso = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).isoformat()
    
    row_c = [{
        'student_id': target_sid,
        'student_name': target_name,
        'course_name': course_name,
        'hours': hours,
        'is_manual': True,
        'timestamp': now_iso
    }]
    errors = bq_client.insert_rows_json(f"{project_id}.grading_data.course_checkins", row_c)
    if errors:
        return jsonify({'success': False, 'error': str(errors)}), 500
    return jsonify({'success': True})

@app.route('/api/admin/attendance_anomalies')
def api_admin_attendance_anomalies():
    if not session.get('is_admin'):
        return jsonify({'success': False, 'error': 'Forbidden'}), 403
    
    try:
        bq_client, project_id = get_bq_client()
        # 查詢遲到 (>08:40), 早退 (<17:00), 或未簽退的紀錄
        query = f"""
            SELECT
              student_name,
              teacher_name,
              sub_room,
              event_date,
              FORMAT_TIMESTAMP('%H:%M', check_in_time, 'Asia/Taipei') as check_in,
              FORMAT_TIMESTAMP('%H:%M', check_out_time, 'Asia/Taipei') as check_out,
              check_in_time,
              check_out_time
            FROM `{project_id}.grading_data.attendance_daily_summary`
            WHERE
                EXTRACT(TIME FROM check_in_time AT TIME ZONE "Asia/Taipei") > TIME(8, 40, 0)
                OR EXTRACT(TIME FROM check_out_time AT TIME ZONE "Asia/Taipei") < TIME(17, 0, 0)
                OR check_out_time IS NULL
            ORDER BY event_date DESC, check_in_time DESC
            LIMIT 100
        """
        results = bq_client.query(query).result()
        data = []
        for r in results:
            data.append({
                'name': r.student_name,
                'teacher': r.teacher_name,
                'room': r.sub_room,
                'date': str(r.event_date),
                'check_in': r.check_in,
                'check_out': r.check_out or '未簽退'
            })
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/pending_epa_feedbacks')
def get_pending_epa_feedbacks():
    is_shared = session.get('is_shared_view', False)
    is_preview = session.get('is_preview_mode', False)
    student_info = session.get('preview_student_info') or session.get('student_info')
    
    if not student_info:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    if not is_shared and not is_preview:
        if not session.get('user'):
            return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    student_id = student_info.get('id')
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        sheet = doc.worksheet('評分記錄')
        all_vals = sheet.get_all_values()
        
        pending = []
        # Header is at index 0, so rows start at index 1
        for idx, r in enumerate(all_vals[1:], start=2):
            # student_id: r[0], comment: r[35], confirmed_time: r[37] (if exists)
            current_sid = str(r[0]).strip()
            if current_sid == student_id:
                # 檢查第 38 欄 (Index 37) 是否已有時間戳記
                confirmed_time = r[37].strip() if len(r) > 37 else ""
                if not confirmed_time:
                    comment = r[35].strip() if len(r) > 35 else ""
                    if comment: # 只有具備質性回饋的才需要確認
                        pending.append({
                            'row_index': idx,
                            'time': r[4] if len(r) > 4 else "未知時間",
                            'station': r[2] if len(r) > 2 else "未知站別",
                            'teacher': r[5] if len(r) > 5 else "未知教師",
                            'comment': comment
                        })
        return jsonify({'success': True, 'data': pending})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/confirm_epa_feedback', methods=['POST'])
def confirm_epa_feedback():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    data = request.json
    row_index = data.get('row_index')
    reply = data.get('reply', '')
    
    if not row_index:
        return jsonify({'success': False, 'error': 'Missing row index'}), 400
        
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        sheet = doc.worksheet('評分記錄')
        
        import datetime
        now_ts = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
        
        # 更新最後兩欄：學員回覆 (37) 與 確認時間 (38)
        # gspread uses 1-indexed for update_cell(row, col, value)
        sheet.update_cell(row_index, 37, reply)
        sheet.update_cell(row_index, 38, now_ts)
        
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/admin/adjust_attendance', methods=['POST'])
def api_admin_adjust_attendance():
    if not session.get('is_admin'):
        return jsonify({'success': False, 'error': 'Forbidden'}), 403
    
    data = request.json
    s_name = data.get('student_name')
    date_str = data.get('date') # YYYY-MM-DD
    room = data.get('room')
    new_in = data.get('new_in')
    new_out = data.get('new_out')
    note = data.get('note', '')
    
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        ws = doc.worksheet('上下班打卡記錄')
        all_rows = ws.get_all_values()
        
        # 尋找匹配行次 (日期、姓名、檢查室)
        target_row_idx = -1
        for i, row in enumerate(all_rows):
            if i == 0: continue
            # 檢查日期 (此處需依據試算表儲存格式靈活比對，通常為 YYYY-MM-DD)
            row_date = row[4].split(' ')[0] if len(row) > 4 else '' # 簽到時間欄在第 5 欄 (index 4)
            if row_date == date_str and row[0] == s_name and row[3] == room:
                target_row_idx = i + 1
                break
        
        if target_row_idx != -1:
            # 更新試算表
            if new_in: ws.update_cell(target_row_idx, 5, f"{date_str} {new_in}:00")
            if new_out: ws.update_cell(target_row_idx, 6, f"{date_str} {new_out}:00")
            if note: ws.update_cell(target_row_idx, 7, note) # 第 7 欄是備註
            
            # TODO: 同步更新 BQ (目前可依賴 sync 腳本或手動更新)
            return jsonify({'success': True})
        else:
            # 若沒找到符合的打卡纪录（例如學員根本沒打卡），但管理員想直接補登請假
            if note == "該生已請假":
                ws.append_row([s_name, "Admin", "", room, date_str, date_str, note])
                return jsonify({'success': True, 'msg': '已補登請假紀錄'})
            
            return jsonify({'success': False, 'error': '找不到對應的出勤紀錄'}), 404
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/admin/sync_bq')
def api_admin_sync_bq():
    """手動觸發 BigQuery 匿名同步"""
    if not session.get('is_admin'):
        return jsonify({'success': False, 'error': 'Forbidden'}), 403
    
    try:
        from sync_to_bq import sync_all
        success = sync_all()
        if success:
            return jsonify({'success': True, 'msg': 'BigQuery 匿名同步完成！'})
        else:
            return jsonify({'success': False, 'error': '同步過程中發生錯誤，請檢查日誌。'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/admin/ai_analysis', methods=['GET', 'POST'])
def admin_ai_analysis():
    if not session.get('user') or not session.get('is_admin'):
        return render_template('error.html', message='權限不足：此頁面僅限系統管理員進入。'), 403
    
    if request.method == 'POST':
        student_name = request.form.get('student_name')
        if not student_name:
            return jsonify({"error": "未提供學員姓名"}), 400
        
        # 調用 AI 分析
        report = generate_ilp_chatgpt(student_name)
        return jsonify({"report": report})
    
    # GET 請求：顯示頁面與學員清單
    def fetch_students():
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        # 嘗試多種可能的分頁名稱
        ws = None
        for name in ['學生名單', '學員名單', '學員名冊']:
            try:
                ws = doc.worksheet(name)
                break
            except: continue
        if not ws: return []
        return safe_get_all_records(ws)

    students = get_cached_data('students', fetch_students)
    return render_template('ai_ilp.html', students=students)

if __name__ == '__main__':
    # 支持 Cloud Run 的動態 Port 綁定，本地預設 5000
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
