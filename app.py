import os
import json
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

def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    # 嘗試多個可能存放憑證的路徑：1. 環境變數 2. Render 預設機密路徑 3. 本機根目錄
    possible_paths = [
        os.getenv("GOOGLE_APPLICATION_CREDENTIALS"),
        os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON"),
        "/etc/secrets/credentials.json",   # Render default secret file path
        "credentials.json"
    ]
    
    creds_file = None
    for path in possible_paths:
        if path and os.path.exists(path):
            creds_file = path
            break

    if creds_file:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
            client = gspread.authorize(creds)
            return client
        except Exception as e:
            print(f"Error loading credentials from {creds_file}: {e}")
            return None
    else:
        print(f"Credentials file not found in any of the configured paths.")
    return None

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
                    student_info = {'id': str(s.get('學生ID', '')), 'name': str(s.get('姓名', ''))}
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
                if alert_needed:
                    threading.Thread(target=send_attendance_alert_email, args=(student_name, teacher_name, sub_room, action, timestamp, time_diff)).start()
                    return jsonify({'success': True, 'msg': f'📤 簽退成功！(⚠️ 系統偵測到早退 {time_diff} 分鐘，已通報)'})
                return jsonify({'success': True, 'msg': '📤 簽退成功！'})
            else:
                # 依指示：若檢查室不一致或找不到紀錄，強迫安插新的一列
                row = [student_name, teacher_name, co_teacher, sub_room, '', timestamp]
                sheet.append_row(row, table_range="A1")
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
