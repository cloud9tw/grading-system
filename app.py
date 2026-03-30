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
    if user:
        return render_template('dashboard.html', user=user)
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
    return redirect('/')

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect('/')

@app.route('/attendance')
def attendance():
    user = session.get('user')
    if not user:
        return redirect('/login')
    return render_template('attendance.html', user=user)

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
        students_records = sheet_students.get_all_records()
        students = [{'id': str(rec.get('學生ID', '')), 'name': str(rec.get('姓名', ''))} for rec in students_records if rec.get('姓名')]
        
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
            teachers_data = sheet_teachers.get_all_records()
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
        
        if not student_name or not sub_room or not action:
            return jsonify({'success': False, 'error': '資料不齊全'}), 400
            
        now_dt = datetime.datetime.utcnow() + datetime.timedelta(hours=8)
        cur_time = now_dt.time()
        
        checkin_std = datetime.time(8, 30, 0)
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
                    
            row = [student_name, teacher_name, sub_room, timestamp, '']
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
                if len(r) >= 3 and str(r[0]).strip() == student_name and str(r[2]).strip() == sub_room:
                    # 如果簽退欄還沒填過
                    if len(r) < 5 or not str(r[4]).strip():
                        found_idx = i
                        break
            
            if found_idx != -1:
                row_num = found_idx + 1
                sheet.update_cell(row_num, 5, timestamp)
                if alert_needed:
                    threading.Thread(target=send_attendance_alert_email, args=(student_name, teacher_name, sub_room, action, timestamp, time_diff)).start()
                    return jsonify({'success': True, 'msg': f'📤 簽退成功！(⚠️ 系統偵測到早退 {time_diff} 分鐘，已通報)'})
                return jsonify({'success': True, 'msg': '📤 簽退成功！'})
            else:
                # 依指示：若檢查室不一致或找不到紀錄，強迫安插新的一列
                row = [student_name, teacher_name, sub_room, '', timestamp]
                sheet.append_row(row, table_range="A1")
                if alert_needed:
                    threading.Thread(target=send_attendance_alert_email, args=(student_name, teacher_name, sub_room, action, timestamp, time_diff)).start()
                    return jsonify({'success': True, 'msg': f'⚠️ 無有效簽到紀錄強制簽退！(且早退 {time_diff} 分，已通報)'})
                return jsonify({'success': True, 'msg': '⚠️ 查無相對應的有效簽到紀錄，已新建獨立新列！'})
            
        else:
            return jsonify({'success': False, 'error': '未知的操作類型。'}), 400
            
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
        students_records = sheet_students.get_all_records()
        students = [{'id': str(rec.get('學生ID', '')), 'name': str(rec.get('姓名', ''))} for rec in students_records if rec.get('姓名')]
        
        # 取得站別OPA細項
        sheet_stations = doc.worksheet('站別OPA細項')
        stations_records = sheet_stations.get_all_records()
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
        trust_records = sheet_trust.get_all_records()
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
            all_records = sheet_records.get_all_records()
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
            teachers_data = sheet_teachers.get_all_records()
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

if __name__ == '__main__':
    # run locally on port 5000
    app.run(debug=True, port=5000)
