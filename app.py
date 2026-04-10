import os
import json
import datetime
from flask import Flask, redirect, url_for, session, render_template, request, jsonify
from authlib.integrations.flask_client import OAuth
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

load_dotenv()

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

def get_bq_client():
    from google.cloud import bigquery
    from google.oauth2 import service_account
    import os
    try:
        credentials = service_account.Credentials.from_service_account_file('credentials.json')
        bq_client = bigquery.Client(credentials=credentials, project=credentials.project_id)
        return bq_client, credentials.project_id
    except Exception as e:
        print("Error initializing BQ Client:", e)
        return None, None

def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    # 1. 直接讀取環境變數中的 JSON 字串 (推薦在 GCP Cloud Run 等無伺服器環境使用)
    json_str = os.getenv("GOOGLE_CREDENTIALS_JSON") or os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if json_str and json_str.strip().startswith('{'):
        try:
            creds_dict = json.loads(json_str)
            return gspread.service_account_from_dict(creds_dict, scopes=scope)
        except Exception as e:
            print(f"Error loading credentials from JSON string: {e}")
            raise Exception(f"GCP JSON credentials parse error: {str(e)}")
            
    # 2. 如果沒有環境變數 JSON，則嘗試讀取實體金鑰檔案
    possible_paths = [
        os.getenv("GOOGLE_APPLICATION_CREDENTIALS"),
        "/etc/secrets/credentials.json",   # Render default secret file path
        "credentials.json"
    ]
    
    errors = []
    for path in possible_paths:
        if path and os.path.exists(path):
            try:
                return gspread.service_account(filename=path, scopes=scope)
            except Exception as e:
                print(f"Error loading credentials from {path}: {e}")
                errors.append(f"[{path}] {str(e)}")
                
    raise Exception("Google Sheets credentials not found. Tried paths: " + " | ".join(errors) if errors else "No valid credential file paths found.")

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

@app.route('/login')
def login():
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
            for t in teachers_records:
                t_email = str(t.get('教師_Email', '')).strip().lower()
                if t_email == user_email:
                    roles.append('teacher')
                    break
                    
            # Check Student
            sheet_students = doc.worksheet('學員名單')
            students_records = safe_get_all_records(sheet_students)
            for s in students_records:
                s_email = str(s.get('Email', '')).strip().lower()
                if s_email == user_email:
                    roles.append('student')
                    student_info = {
                        'id': str(s.get('學生ID', '')),
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
            # AI: 未登錄教師姓名 (重複 F 欄)
            data.get('other_teacher', '')
        ]
        # 找到最後一行有資料的位置，再往下一行寫入
        all_values = sheet.get_all_values()
        next_row = len(all_values) + 1  # 從第 1 行算起，下一個空行
        sheet.update(f'A{next_row}', [row])
        
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
            
        gc = get_gspread_client()
        if not gc:
            return jsonify({'success': False, 'error': 'No gspread client'})
            
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        # 取得所有學員
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
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        sheet = doc.worksheet('評分記錄')
        all_vals = sheet.get_all_values()
        
        records_by_station = {}
        target_id = student_info['id']
        
        for r in all_vals[1:]:
            if len(r) > 8 and str(r[0]).strip() == target_id:
                station = str(r[2]).strip()
                time_str = str(r[4]).strip()
                opa1 = str(r[6]).strip()
                opa2 = str(r[7]).strip()
                opa3 = str(r[8]).strip()
                
                def to_int(v):
                    try: return int(v)
                    except: return None
                        
                pt = {
                    'time': time_str,
                    'opa1': to_int(opa1),
                    'opa2': to_int(opa2),
                    'opa3': to_int(opa3)
                }
                
                if station not in records_by_station:
                    records_by_station[station] = []
                records_by_station[station].append(pt)
                
        for st in records_by_station:
            records_by_station[st].sort(key=lambda x: x['time'])
            
        return jsonify({'success': True, 'data': records_by_station})
    except Exception as e:
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
    if session.get('current_role') != 'student':
        return jsonify({'success': False, 'error': 'Forbidden'}), 403
        
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
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        sheet = doc.worksheet('上下班打卡記錄')
        all_vals = sheet.get_all_values()
        
        target_name = student_info['name']
        history = []
        
        for r in all_vals[1:]:
            # Header: ['學生', '教師', '共同教師', '檢查室', '簽到時間', '簽退時間']
            if len(r) >= 5 and str(r[0]).strip() == target_name:
                history.append({
                    'teacher': str(r[1]).strip(),
                    'co_teacher': str(r[2]).strip() if len(r) > 2 else '',
                    'room': str(r[3]).strip() if len(r) > 3 else '',
                    'check_in': str(r[4]).strip() if len(r) > 4 else '',
                    'check_out': str(r[5]).strip() if len(r) > 5 else ''
                })
        
        # Sort by check_in time descending (newest first)
        history.sort(key=lambda x: x['check_in'], reverse=True)
                
        return jsonify({'success': True, 'data': history})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    # run locally on port 5000
    app.run(debug=True, port=5000)
