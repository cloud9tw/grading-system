import datetime
import re
import os
from google.cloud import bigquery
from google.oauth2 import service_account

def get_first_monday():
    return datetime.date(2025, 7, 7)

def parse_exemptions(doc):
    try:
        sheet = doc.worksheet('排除計時日期')
        vals = sheet.get_all_values()
        exemptions = {}
        for r in vals[1:]:
            if r and r[0]:
                date_str = str(r[0]).strip()
                try:
                    date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                    accounts = str(r[2]).strip() if len(r) > 2 else ""
                    exemptions[date_obj] = [ac.strip() for ac in accounts.split(',')] if accounts else []
                except: pass
        return exemptions
    except:
        return {}

def process_student_gamification(s_id, s_name, s_type, s_gender, s_email, epa_vals, grade_vals, room_vals, fb_vals, attendance_vals, exemptions, first_monday, today):
    # Standardize s_id and s_name
    s_id = str(s_id).strip()
    s_name = str(s_name).strip()

    # === 1. EPA Progress ===
    epa_progress = {}
    type_idx = -1
    if epa_vals and len(epa_vals) > 0:
        header = epa_vals[0]
        for i, h in enumerate(header):
            if h.strip() == s_type:
                type_idx = i; break
    
    curr_cat = "未分類"
    if type_idx != -1:
        for r in epa_vals[1:]:
            if not r or not r[0]: continue
            station = str(r[0]).strip()
            if station.upper().startswith('EPA'):
                curr_cat = station; continue
            # Check if this station has a target for this student type
            val_str = str(r[type_idx]) if type_idx < len(r) else ''
            m = re.search(r'\d+', val_str)
            if m:
                if curr_cat not in epa_progress: epa_progress[curr_cat] = {}
                epa_progress[curr_cat][station] = {'target': int(m.group(0)), 'current': 0}

    # Count Progress from Grading Logs (ID based primary, Name fallback)
    for r in grade_vals[1:]:
        if len(r) > 3:
            r_id = str(r[0]).strip()
            r_name = str(r[1]).strip()
            # Match by clean ID or Name
            if (s_id and r_id == s_id) or (r_name == s_name):
                station_dept = str(r[2]).strip()
                body_part = str(r[3]).strip()
                for cat, stations in epa_progress.items():
                    for epa_k, prog in stations.items():
                        # Matching logic: if station or body part is in the target key
                        if body_part in epa_k or epa_k in body_part or epa_k in station_dept:
                            prog['current'] += 1
                            break

    # === 2. Feedback Progress ===
    feedback_progress = {}
    if s_type == '實習學生':
        for r in room_vals:
            if not r: continue
            dept = str(r[0]).strip()
            if not dept: continue
            m = re.search(r'\d+', str(r[-1]).strip())
            if m:
                target_val = int(m.group(0))
                if 'Mammo' in dept and s_gender == '男': target_val = 0
                if target_val > 0: feedback_progress[dept] = {'target': target_val, 'current': 0}
        for r in fb_vals[1:]:
            # Use Name for feedback since BQ doesn't have Student ID in feedback sheet yet
            if len(r) > 2 and str(r[2]).strip() == s_name:
                dept = str(r[6]).strip()
                found_match = False
                if '急診' in dept:
                    for k in feedback_progress.keys():
                        if 'Routine' in k:
                            feedback_progress[k]['current'] += 1; found_match = True; break
                if not found_match:
                    for k in feedback_progress.keys():
                        if k in dept or dept in k:
                            feedback_progress[k]['current'] += 1; break

    # === 3. Points & Medals Calculation ===
    points = 0
    medals = []
    
    epa_master = True
    has_epa_targets = False
    for cat, stations in epa_progress.items():
        cat_done = True
        cat_has_data = False
        for k, v in stations.items():
            cat_has_data = True
            has_epa_targets = True
            points += v['current'] * 10
            if v['current'] < v['target']:
                cat_done = False
                epa_master = False
        if cat_has_data:
            name_prefix = cat.split('-')[0]
            medals.append({'name': f"{name_prefix}專精", 'desc': f"完成 {cat} 所有需求", 'achieved': cat_done})
            
    if has_epa_targets:
        medals.append({'name': 'EPA 大師', 'desc': '所有身分設定之 EPA 分類全數解鎖', 'achieved': epa_master})
        
    feedback_master = True
    has_fb_targets = False
    for k, v in feedback_progress.items():
        has_fb_targets = True
        points += v['current'] * 5
        fb_done = (v['current'] >= v['target'])
        if not fb_done: feedback_master = False
        name_prefix = k.split('(')[-1].replace(')', '') if '(' in k else k
        medals.append({'name': f"{name_prefix}回饋召集", 'desc': f"繳交 {k} 教學回饋表單", 'achieved': fb_done})
        
    if has_fb_targets:
        medals.append({'name': '回饋達人', 'desc': '所有目標站別之回饋表單全數達成', 'achieved': feedback_master})
        
    # === 4. Attendance ===
    attended_dates = {}
    for r in attendance_vals[1:]:
        if len(r) > 5 and str(r[0]).strip() == s_name:
            # Timestamp format in BQ: 2026-04-10T08:46:43
            try:
                cin = str(r[5]).replace('T', ' ').split('.')[0]
                d = datetime.datetime.strptime(cin, "%Y-%m-%d %H:%M:%S").date()
                if d not in attended_dates: attended_dates[d] = {'in': True, 'out': True} # BQ daily summary already implies both if we use that view
            except: pass
                    
    # (Simplified attendance logic for performance in leaderboard)
    perfect_months = 0
    # ... (skipping some complex perfect month logic for leaderboard brevity if needed, but keeping for student)
    
    return {
        'points': points,
        'epa_progress': epa_progress,
        'feedback_progress': feedback_progress,
        'medals': medals,
        'achieved_count': sum(1 for m in medals if m['achieved'])
    }

def get_bq_gamification_logs():
    try:
        credentials = service_account.Credentials.from_service_account_file('credentials.json')
        client = bigquery.Client(credentials=credentials, project=credentials.project_id)
        project = credentials.project_id
        
        # 1. Grading
        q_gps = f"SELECT student_id, student_name, station, body_part, timestamp, teacher_name, opa1_sum, opa2_sum, opa3_sum, opa1_items, opa2_items, opa3_items, aspect1, aspect2, comment FROM `{project}.grading_data.grading_logs` WHERE is_deleted = FALSE OR is_deleted IS NULL"
        res_gps = client.query(q_gps).result()
        gps = [['ID', 'Name', 'Station', 'BodyPart', 'Time', 'Teacher', 'OPA1', 'OPA2', 'OPA3']]
        for r in res_gps:
            row = [r.student_id, r.student_name, r.station, r.body_part, r.timestamp.isoformat() if r.timestamp else '', r.teacher_name, r.opa1_sum, r.opa2_sum, r.opa3_sum]
            gps.append(row)
            
        # 2. Feedback
        q_fps = f"SELECT student_name, department, timestamp FROM `{project}.grading_data.feedback_logs` WHERE is_deleted = FALSE OR is_deleted IS NULL"
        res_fps = client.query(q_fps).result()
        fps = [['', '', 'Name', '', '', '', 'Dept']]
        for r in res_fps:
            fps.append(['', '', r.student_name, '', '', '', r.department])
            
        # 3. Attendance
        q_aps = f"SELECT student_name, event_time as event_time FROM `{project}.grading_data.attendance_events` WHERE is_deleted = FALSE OR is_deleted IS NULL"
        res_aps = client.query(q_aps).result()
        aps = [['Name', '', '', '', 'Time', 'Time']]
        for r in res_aps:
            aps.append([r.student_name, '', '', '', r.event_time.isoformat(), r.event_time.isoformat()])
            
        return gps, fps, aps
    except Exception as e:
        print("BQ Gamification fetch error:", e)
        return [], [], []

def get_student_gamification_data(gc, doc, student_info):
    eps = doc.worksheet('各類別EPA需求').get_all_values()
    rps = doc.worksheet('檢查室清單').get_all_values()
    gps, fps, aps = get_bq_gamification_logs()
    exps = parse_exemptions(doc)
    
    s_email = ""
    st_vals = doc.worksheet('學員名單').get_all_values()
    for r in st_vals[1:]:
        if len(r) > 2 and str(r[0]).strip() == str(student_info.get('id', '')).strip():
            s_email = str(r[2]).strip()
            break
            
    return process_student_gamification(
        student_info.get('id', ''), student_info.get('name', ''), student_info.get('type', ''), student_info.get('gender', ''), s_email,
        eps, gps, rps, fps, aps, exps, get_first_monday(), datetime.date.today()
    )

def get_leaderboard_data(gc, doc):
    try:
        eps = doc.worksheet('各類別EPA需求').get_all_values()
        rps = doc.worksheet('檢查室清單').get_all_values()
        exps = parse_exemptions(doc)
        fm = get_first_monday()
        today = datetime.date.today()
        
        st_vals = doc.worksheet('學員名單').get_all_values()
        header = st_vals[0]
        id_idx = header.index('學生ID') if '學生ID' in header else 0
        name_idx = header.index('姓名') if '姓名' in header else 1
        role_idx = header.index('職級') if '職級' in header else 4
        
        gps, fps, aps = get_bq_gamification_logs()
        
        leaderboard = []
        for r in st_vals[1:]:
            if len(r) > role_idx and str(r[role_idx]).strip() == '實習學生':
                v = {
                    'id': str(r[id_idx]).strip(),
                    'name': str(r[name_idx]).strip(),
                    'type': '實習學生',
                    'gender': str(r[header.index('性別')]).strip() if '性別' in header else '',
                    'email': str(r[2]).strip() if len(r) > 2 else ''
                }
                data = process_student_gamification(v['id'], v['name'], v['type'], v['gender'], v['email'],
                                                   eps, gps, rps, fps, aps, exps, fm, today)
                leaderboard.append({
                    'name': v['name'],
                    'points': data['points'],
                    'medals': data['achieved_count']
                })
        
        leaderboard.sort(key=lambda x: x['points'], reverse=True)
        return leaderboard
    except Exception as e:
        import traceback
        traceback.print_exc()
        return []
