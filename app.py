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
    creds_file = os.getenv("GOOGLE_APPLICATION_CREDENTIALS") or os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if creds_file and os.path.exists(creds_file):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
            client = gspread.authorize(creds)
            return client
        except Exception as e:
            print(f"Error loading credentials: {e}")
            return None
    else:
        print(f"Credentials file not found or path not set: {creds_file}")
    return None

@app.route('/')
def index():
    user = session.get('user')
    if user:
        return render_template('dashboard.html', user=user)
    return render_template('login.html')

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
            station_data = {
                'name': name,
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
            trust_levels.append({
                'score': str(rec.get('分數', '')),
                'level': str(rec.get('信賴等級', '')),
                'desc': str(rec.get('描述', ''))
            })
            
        # 取得歷史評分紀錄次數
        student_stats = {}
        try:
            sheet_records = doc.worksheet('評分記錄')
            # 取得所有評分紀錄
            all_records = sheet_records.get_all_records()
            for rec in all_records:
                sid = str(rec.get('ID', '')).strip()
                stn = str(rec.get('站別', '')).strip()
                if sid and stn:
                    if sid not in student_stats:
                        student_stats[sid] = {'stations': {}, 'aspects': {}}
                    if stn not in student_stats[sid]['stations']:
                        student_stats[sid]['stations'][stn] = 0
                    student_stats[sid]['stations'][stn] += 1
                    
                    # 尋找所有名稱內包含 '面向選擇' 的欄位並統計其次數
                    for k, v in rec.items():
                        if '面向選擇' in str(k) and str(v).strip():
                            import re
                            # 擷取開頭的數字作為代號
                            m = re.match(r'^\d+', str(v).strip())
                            if m:
                                asp_num = m.group()
                                if asp_num not in student_stats[sid]['aspects']:
                                    student_stats[sid]['aspects'][asp_num] = 0
                                student_stats[sid]['aspects'][asp_num] += 1
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
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
        
        # 收集 34 欄資料
        student_id = data.get('student_id', '')
        student_name = data.get('student_name', '')
        station = data.get('station', '')
        
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
        
        sheet.append_row(row)
        
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    # run locally on port 5000
    app.run(debug=True, port=5000)
