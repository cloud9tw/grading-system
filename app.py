import os
import json
import datetime
import logging
from google.cloud import bigquery
from google.oauth2 import service_account
from flask import Flask, redirect, url_for, session, render_template, request, jsonify, Response
from authlib.integrations.flask_client import OAuth
import gspread
from dotenv import load_dotenv
import threading
from ceep_scraper import scrape_ceep_all_forms
from ceep_archiver import archive_to_sheets
from sync_to_bq import sync_all as sync_to_bq_all

load_dotenv()
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "super-secret-key-12345")

from werkzeug.middleware.proxy_fix import ProxyFix
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

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
            student_info = session.get('student_info', {})
            s_id = student_info.get('id', '')
            root_url = request.url_root
            target_url = f"{root_url}attendance?student_id={s_id}"
            qr_img_url = f"https://api.qrserver.com/v1/create-qr-code/?size=250x250&data={urllib.parse.quote(target_url)}"
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
        return render_template('error.html', message='教學回饋表僅限學員身份使用。若您同時具有教師與學員身份，請先切換至「學員介面」再團入。'), 403
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
        now_dt = datetime.datetime.utcnow() + datetime.timedelta(hours=8)
        if now_dt.weekday() >= 5: # 週末不寄信
            return jsonify({'success': True, 'msg': 'Today is weekend, skip absent check.'})
            
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        # 1. 檢查排除日期 (系統選定排除日期，例如國定假日)
        try:
            sheet_settings = doc.worksheet('系統設定')
            # 排除日期在第二欄 (B欄)
            excluded_dates = sheet_settings.col_values(2)[1:] # 跳過標題
            if today_str in [str(d).strip() for d in excluded_dates if d]:
                return jsonify({'success': True, 'msg': f'Today ({today_str}) is an excluded date, skip absent check.'})
        except Exception as e:
            print(f"Read excluded dates error: {e}")

        # 2. 取得所有學員
        sheet_students = doc.worksheet('學員名單')
        students_records = safe_get_all_records(sheet_students)
        all_students = [str(rec.get('姓名', '')).strip() for rec in students_records if str(rec.get('姓名', '')).strip()]
        if not all_students:
            return jsonify({'success': True, 'msg': 'No students in roster'})
        
        # 取得今日打卡
        today_str = now_dt.strftime("%Y-%m-%d")
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
        
        # 取得學員名單
        sheet_students = doc.worksheet('學員名單')
        # [學生ID, 姓名, email]
        students_records = safe_get_all_records(sheet_students)
        students = [{'id': str(rec.get('學生ID', '')), 'name': str(rec.get('姓名', '')), 'email': str(rec.get('Email', '')).strip().lower()} for rec in students_records if rec.get('姓名')]
        
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
            
        # 取得歷史評分紀錄次數
        student_stats = {}
        try:
            sheet_records = doc.worksheet('評分記錄')
            # 取得所有評分紀錄
            all_records = safe_get_all_records(sheet_records)
            for rec in all_records:
                sid = str(rec.get('學員ID', '')).strip()
                if not sid:
                    sid = str(rec.get('ID', '')).strip()
                
                stn = str(rec.get('站別', '')).strip()
                bpart = str(rec.get('檢查部位', '')).strip()
                
                if sid and stn:
                    if sid not in student_stats:
                        student_stats[sid] = {'stations': {}}
                    if stn not in student_stats[sid]['stations']:
                        student_stats[sid]['stations'][stn] = {
                            'count': 0,
                            'body_parts': {},
                            'aspects': {}
                        }
                    
                    student_stats[sid]['stations'][stn]['count'] += 1
                    
                    if bpart:
                        if bpart not in student_stats[sid]['stations'][stn]['body_parts']:
                            student_stats[sid]['stations'][stn]['body_parts'][bpart] = 0
                        student_stats[sid]['stations'][stn]['body_parts'][bpart] += 1
                    
                    for k, v in rec.items():
                        if '面向選擇' in str(k) and str(v).strip():
                            import re
                            m = re.match(r'^\d+', str(v).strip())
                            if m:
                                asp_num = m.group()
                                if asp_num not in student_stats[sid]['stations'][stn]['aspects']:
                                    student_stats[sid]['stations'][stn]['aspects'][asp_num] = 0
                                student_stats[sid]['stations'][stn]['aspects'][asp_num] += 1
        except Exception as sheet_err:
            print("Error parsing 評分記錄 for stats:", sheet_err)
            
        return jsonify({
            'success': True,
            'students': students,
            'stations': stations,
            'trust_levels': trust_levels,
            'student_stats': student_stats
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/submit_grade', methods=['POST'])
def submit_grade():
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
        sheet = doc.worksheet('評分記錄')
        
        import datetime
        timestamp = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
        teacher_email = user.get('email', '')
        teacher_name = teacher_email # default
        
        # 查詢教師名單中的姓名
        try:
            sheet_teachers = doc.worksheet('教師名單')
            teachers_data = safe_get_all_records(sheet_teachers)
            for rec in teachers_data:
                # 欄位： 教師_Email, 教師姓名
                t_email = str(rec.get('教師_Email', '')).strip().lower()
                if t_email == teacher_email.lower():
                    n = str(rec.get('教師姓名', '')).strip()
                    if n:
                        teacher_name = n
                    break
        except Exception as e:
            print("Error loading teacher names:", e)
        
        # 收集資料
        student_id = data.get('student_id', '')
        student_name = data.get('student_name', '')
        station = data.get('station', '')
        body_part = data.get('body_part', '')
        
        # 總和評比
        opa1_sum = data.get('opa1_sum', '')
        opa2_sum = data.get('opa2_sum', '')
        opa3_sum = data.get('opa3_sum', '')
        
        # 細項 1~8 (陣列)
        opa1_items = data.get('opa1_items', [''] * 8)
        opa2_items = data.get('opa2_items', [''] * 8)
        opa3_items = data.get('opa3_items', [''] * 8)
        
        # 如果長度不足 8，補齊空白
        opa1_items = (opa1_items + [''] * 8)[:8]
        opa2_items = (opa2_items + [''] * 8)[:8]
        opa3_items = (opa3_items + [''] * 8)[:8]
        
        aspect1 = data.get('aspect1', '')
        aspect2 = data.get('aspect2', '')
        comment = data.get('comment', '')
        
        row = [
            student_id,
            student_name,
            station,
            body_part,
            timestamp,
            teacher_name,
            opa1_sum,
            opa2_sum,
            opa3_sum
        ] + opa1_items + opa2_items + opa3_items + [
            aspect1,
            aspect2,
            comment
        ]
        # 強制告訴 Google 試算表從整張表的 A1 欄開始往下當作新增基準
        # 避免 Google 的「自動探測表格」功能發生向右位移（例如跳過前 32 欄）的 BUG
        sheet.append_row(row, table_range="A1")
        
        # BQ Double-Write
        try:
            bq_client, project_id = get_bq_client()
            if bq_client:
                # 重新解析 ISO Timestamp
                dt_iso = datetime.datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S").isoformat()
                grade_insert = [{
                    'student_id': student_id, 'student_name': student_name, 'station': station,
                    'body_part': body_part, 'timestamp': dt_iso, 'teacher_name': teacher_name,
                    'opa1_sum': opa1_sum, 'opa2_sum': opa2_sum, 'opa3_sum': opa3_sum,
                    'opa1_items': opa1_items, 'opa2_items': opa2_items, 'opa3_items': opa3_items,
                    'aspect1': aspect1, 'aspect2': aspect2, 'comment': comment
                }]
                errors = bq_client.insert_rows_json(f"{project_id}.grading_data.grading_logs", grade_insert)
                if errors: print("BQ Grading Insert Error:", errors)
        except Exception as e:
            print("BQ Insert Error (Grading):", e)
            
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/student_stats', methods=['GET'])
def get_student_stats():
    user = session.get('user')
    if not user:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
        
    if session.get('current_role') != 'student':
        return jsonify({'success': False, 'error': 'Forbidden: Not acting as student'}), 403
        
    student_info = session.get('student_info')
    if not student_info:
        return jsonify({'success': False, 'error': 'No student info found.'}), 400
        
    try:
        # Step 1: Normalize Target ID from session (Ensure it's a clean string)
        target_id = str(student_info.get('id', '')).split('.')[0].strip()
        target_name = str(student_info.get('name', '')).strip()
        
        import logging
        logging.info(f"API: [Stats] Fetching for ID: [{target_id}], Name: [{target_name}]")
        
        records_by_station = {}
        bq_client, project_id = get_bq_client()
        
        if bq_client:
            # Use CAST in SQL to handle cases where student_id might be INTEGER in BQ vs STRING in Session
            q = f"""
                SELECT station, timestamp, opa1_sum, opa2_sum, opa3_sum, teacher_name, aspect1, aspect2, comment
                FROM `{project_id}.grading_data.grading_logs` 
                WHERE (CAST(student_id AS STRING) = @sid OR student_name = @sname) 
                AND (is_deleted = FALSE OR is_deleted IS NULL) 
                ORDER BY timestamp ASC
            """
            job_config = bigquery.QueryJobConfig(
                query_parameters=[
                    bigquery.ScalarQueryParameter("sid", "STRING", target_id),
                    bigquery.ScalarQueryParameter("sname", "STRING", target_name)
                ]
            )
            query_job = bq_client.query(q, job_config=job_config)
            res = query_job.result(timeout=20) 
            
            row_count = 0
            def to_int(v):
                try: return int(v)
                except: return None
                
            for r in res:
                row_count += 1
                station = str(r.station or 'Unknown').strip()
                if station not in records_by_station:
                    records_by_station[station] = []
                records_by_station[station].append({
                    'time': r.timestamp.strftime('%Y/%m/%d %H:%M:%S') if r.timestamp else '',
                    'teacher': r.teacher_name or '未知',
                    'opa1': to_int(r.opa1_sum),
                    'opa2': to_int(r.opa2_sum),
                    'opa3': to_int(r.opa3_sum),
                    'aspect1': r.aspect1 or '',
                    'aspect2': r.aspect2 or '',
                    'comment': r.comment or ''
                })
            logging.info(f"API: Found {row_count} rows for student_id: [{target_id}] in project: [{project_id}]")
            
            # Step 2: Fetch Course Check-in Hours
            course_hours = 0.0
            try:
                q_course = f"""
                    SELECT SUM(hours) as total 
                    FROM `{project_id}.grading_data.course_checkins` 
                    WHERE CAST(student_id AS STRING) = @sid
                """
                job_config_c = bigquery.QueryJobConfig(
                    query_parameters=[bigquery.ScalarQueryParameter("sid", "STRING", target_id)]
                )
                res_c = bq_client.query(q_course, job_config=job_config_c).result()
                for r_c in res_c:
                    course_hours = float(r_c.total or 0.0)
            except Exception as e:
                logging.error(f"Error fetching course hours: {e}")

            return jsonify({
                'success': True, 
                'data': records_by_station, 
                'count': row_count,
                'course_hours': course_hours
            })
        else:
            logging.error("API: BigQuery client could not be initialized (likely missing credentials.json)")
            return jsonify({'success': False, 'error': 'BigQuery client missing'}), 500
    except Exception as e:
        import traceback
        traceback.print_exc()
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
        from gamification import get_student_gamification_data
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        data = get_student_gamification_data(gc, doc, student_info)
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
        from gamification import get_leaderboard_data
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        data = get_leaderboard_data(gc, doc)
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
        # Prepare name variations to avoid REPLACE in SQL
        name_orig = student_info['name']
        name_clean = name_orig.replace(' ', '').replace('　', '')
        search_names = list(set([name_orig, name_clean]))
        
        history = []
        bq_client, project_id = get_bq_client()
        if bq_client:
            q = f"SELECT sub_room, check_in_time, check_out_time, teacher_name, co_teacher FROM `{project_id}.grading_data.attendance_daily_summary` WHERE student_name IN UNNEST(@names) ORDER BY check_in_time DESC"
            job_config = bigquery.QueryJobConfig(
                query_parameters=[bigquery.ArrayQueryParameter("names", "STRING", search_names)]
            )
            # Add explicit timeout of 20 seconds to prevent hanging the Flask thread
            query_job = bq_client.query(q, job_config=job_config)
            res = query_job.result(timeout=20)
            
            for r in res:
                # Convert row to dict for safer access across different versions of BQ client
                row_dict = dict(r)
                history.append({
                    'room': row_dict.get('sub_room', ''),
                    'check_in': row_dict.get('check_in_time').strftime('%Y/%m/%d %H:%M:%S') if row_dict.get('check_in_time') else '',
                    'check_out': row_dict.get('check_out_time').strftime('%Y/%m/%d %H:%M:%S') if row_dict.get('check_out_time') else '',
                    'teacher': row_dict.get('teacher_name', ''),
                    'co_teacher': row_dict.get('co_teacher', ''),
                    'is_complete': bool(row_dict.get('check_in_time') and row_dict.get('check_out_time'))
                })
        
        return jsonify({'success': True, 'data': history})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# --- 管理員門戶與功能 ---

@app.route('/admin')
def admin_portal():
    user = session.get('user')
    if not user or not session.get('is_admin'):
        return render_template('error.html', message='權限不足：此頁面僅限系統管理員進入。'), 403
    
    # 傳遞 Sheets ID 供前端產生連結
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    sheet_feedback = "112l_e3WKbIkFYj58nv8LRTYEvfyDpXMh-NcSe98T07w"
    
    # 取得學員名單供模擬功能使用
    students = []
    try:
        gc = get_gspread_client()
        doc = gc.open_by_key(sheet_id)
        sheet_students = doc.worksheet('學員名單')
        records = safe_get_all_records(sheet_students)
        for r in records:
            s_name = str(r.get('姓名', '')).strip()
            s_id = str(r.get('學生ID', '')).strip()
            if s_name and s_id:
                students.append({'name': s_name, 'id': s_id})
    except Exception as e:
        print(f"Admin fetch students error: {e}")

    return render_template('admin_portal.html', 
                           user=user, 
                           roles=session.get('roles', []),
                           sheet_main=sheet_id,
                           sheet_feedback=sheet_feedback,
                           students=students)

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
    user = session.get('user')
    student_info = session.get('student_info')
    if not user or not student_info:
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


if __name__ == '__main__':
    # run locally on port 5000
    app.run(debug=True, port=5000)
