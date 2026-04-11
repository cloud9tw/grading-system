import datetime
import re
import os
from google.cloud import bigquery
from google.oauth2 import service_account

def get_first_monday():
    return datetime.date(2025, 7, 7)

def parse_scoring_config(doc):
    # Default values
    config = {
        'epa_score': 10,
        'fb_score': 5,
        'epa_bonus': 100,
        'fb_bonus': 50
    }
    try:
        sheet = doc.worksheet('遊戲化配分設定')
        vals = sheet.get_all_values()
        for r in vals[1:]:
            if len(r) >= 2:
                key = str(r[0]).strip()
                try:
                    val = int(r[1])
                    if key == 'EPA評核得分': config['epa_score'] = val
                    elif key == '教學回饋得分': config['fb_score'] = val
                    elif key == '達成所有EPA加分': config['epa_bonus'] = val
                    elif key == '達成所有回饋加分': config['fb_bonus'] = val
                except: pass
    except:
        pass
    return config

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

def group_logs_by_student(grade_vals, fb_vals):
    # Mapping standardized name/id to their records
    indexed_grades = {}
    indexed_fb = {}
    
    # Process grading logs
    if grade_vals and len(grade_vals) > 1:
        for r in grade_vals[1:]:
            if len(r) > 3:
                r_id = str(r[0]).strip()
                r_name = str(r[1]).strip().replace(' ', '').replace('　', '')
                
                # Create a list for both the ID and the clean Name
                if r_id:
                    # Handle both integer and float string IDs
                    base_id = r_id.split('.')[0]
                    if base_id not in indexed_grades: indexed_grades[base_id] = []
                    indexed_grades[base_id].append(r)
                if r_name:
                    if r_name not in indexed_grades: indexed_grades[r_name] = []
                    indexed_grades[r_name].append(r)
    
    # Process feedback logs
    if fb_vals and len(fb_vals) > 1:
        for r in fb_vals[1:]:
            if len(r) > 2:
                r_name = str(r[2]).strip().replace(' ', '').replace('　', '')
                if r_name:
                    if r_name not in indexed_fb: indexed_fb[r_name] = []
                    indexed_fb[r_name].append(r)
                    
    return indexed_grades, indexed_fb

def process_student_gamification(s_id, s_name, s_type, s_gender, s_email, epa_vals, grade_vals, room_vals, fb_vals, attendance_vals, exemptions, first_monday, today, scoring_config=None, indexed_data=None):
    if scoring_config is None:
        scoring_config = {'epa_score': 10, 'fb_score': 5, 'epa_bonus': 100, 'fb_bonus': 50}
        
    # Standardize s_id and s_name
    s_id = str(s_id).strip()
    s_name = str(s_name).strip()
    clean_s_name = s_name.replace(' ', '').replace('　', '') # Remove both half-width and full-width spaces

    # Normalize Student Type
    if not s_type or s_type.strip() == "":
        s_type = '實習學生'

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
            val_str = str(r[type_idx]) if type_idx < len(r) else ''
            m = re.search(r'\d+', val_str)
            if m:
                if curr_cat not in epa_progress: epa_progress[curr_cat] = {}
                epa_progress[curr_cat][station] = {'target': int(m.group(0)), 'current': 0}

    # Count EPA Progress using either indexed data or scanning the whole list
    relevant_grades = []
    if indexed_data and 'grades' in indexed_data:
        # Check both ID and Name in the index
        relevant_grades = indexed_data['grades'].get(s_id, [])
        if not relevant_grades:
            relevant_grades = indexed_data['grades'].get(clean_s_name, [])
    else:
        # Fallback to scanning (Keep for individual calls)
        for r in grade_vals[1:]:
            if len(r) > 3:
                r_id = str(r[0]).strip()
                r_name = str(r[1]).strip().replace(' ', '').replace('　', '')
                if (s_id and (r_id == s_id or r_id == s_id + ".0")) or (r_name == clean_s_name):
                    relevant_grades.append(r)

    for r in relevant_grades:
        # Handle both raw sheet rows (list) and aggregate BQ data (dict)
        station_dept = ""
        body_part = ""
        count_to_add = 1
        
        if isinstance(r, dict):
            station_dept = str(r.get('station', '')).strip()
            body_part = str(r.get('body_part', '')).strip()
            count_to_add = int(r.get('cnt', 1))
        else:
            station_dept = str(r[2]).strip()
            body_part = str(r[3]).strip()
            count_to_add = 1

        for cat, stations in epa_progress.items():
            for epa_k, prog in stations.items():
                if body_part in epa_k or epa_k in body_part or epa_k in station_dept:
                    prog['current'] += count_to_add
                    break

    # === 2. Feedback Progress ===
    feedback_progress = {}
    for r in room_vals:
        if not r: continue
        dept = str(r[0]).strip()
        if not dept: continue
        m = re.search(r'\d+', str(r[-1]).strip())
        if m:
            target_val = int(m.group(0))
            if 'Mammo' in dept and s_gender == '男': target_val = 0
            if target_val > 0: feedback_progress[dept] = {'target': target_val, 'current': 0}
            
    relevant_fb = []
    if indexed_data and 'fb' in indexed_data:
        relevant_fb = indexed_data['fb'].get(clean_s_name, [])
    else:
        for r in fb_vals[1:]:
            if len(r) > 2:
                r_name = str(r[2]).strip().replace(' ', '').replace('　', '')
                if r_name == clean_s_name:
                    relevant_fb.append(r)

    for r in relevant_fb:
        dept = ""
        count_to_add = 1
        
        if isinstance(r, dict):
            dept = str(r.get('dept', '')).strip()
            count_to_add = int(r.get('cnt', 1))
        else:
            dept = str(r[6]).strip()
            count_to_add = 1
            
        found_match = False
        if '急診' in dept:
            for k in feedback_progress.keys():
                if 'Routine' in k:
                    feedback_progress[k]['current'] += count_to_add; found_match = True; break
        if not found_match:
            for k in feedback_progress.keys():
                if k in dept or dept in k:
                    feedback_progress[k]['current'] += count_to_add; break

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
            points += v['current'] * scoring_config['epa_score']
            if v['current'] < v['target']:
                cat_done = False
                epa_master = False
        if cat_has_data:
            name_prefix = cat.split('-')[0]
            medals.append({'name': f"{name_prefix}專精", 'desc': f"完成 {cat} 所有需求", 'achieved': cat_done})
            
    if has_epa_targets:
        if epa_master: points += scoring_config['epa_bonus']
        medals.append({'name': 'EPA 大師', 'desc': '所有身分設定之 EPA 分類全數解鎖', 'achieved': epa_master})
        
    feedback_master = True
    has_fb_targets = False
    for k, v in feedback_progress.items():
        has_fb_targets = True
        points += v['current'] * scoring_config['fb_score']
        fb_done = (v['current'] >= v['target'])
        if not fb_done: feedback_master = False
        name_prefix = k.split('(')[-1].replace(')', '') if '(' in k else k
        medals.append({'name': f"{name_prefix}回饋召集", 'desc': f"繳交 {k} 教學回饋表單", 'achieved': fb_done})
        
    if has_fb_targets:
        if feedback_master: points += scoring_config['fb_bonus']
        medals.append({'name': '回饋達人', 'desc': '所有目標站別之回饋表單全數達成', 'achieved': feedback_master})
        
    # === 3.5 Internship Week ===
    internship_week = 0
    if first_monday and today:
        delta = today - first_monday
        internship_week = (delta.days // 7) + 1
        
    # === 4. Return ===
    return {
        'points': points,
        'epa_progress': epa_progress,
        'feedback_progress': feedback_progress,
        'medals': medals,
        'achieved_count': sum(1 for m in medals if m['achieved']),
        'student_type': s_type,
        'internship_week': internship_week
    }

def get_bq_gamification_logs():
    """
    Returns aggregated counts from BigQuery to minimize data transfer and memory usage.
    Returns: (grading_counts, feedback_counts)
    """
    try:
        try:
            credentials = service_account.Credentials.from_service_account_file('credentials.json')
            client = bigquery.Client(credentials=credentials, project=credentials.project_id)
            project = credentials.project_id
        except FileNotFoundError:
            import google.auth
            credentials, project = google.auth.default()
            project = project or "epa-grading-system"
            client = bigquery.Client(credentials=credentials, project=project)
        
        # 1. Aggregated Grading Logs: Count per (student_id, student_name, station, body_part)
        # Using CAST to ensure student_id is always a string matching our session '8' vs '8.0'
        q_gps = f"""
            SELECT CAST(student_id AS STRING) as sid, student_name as sname, 
                   station, body_part, COUNT(*) as cnt 
            FROM `{project}.grading_data.grading_logs` 
            WHERE is_deleted = FALSE OR is_deleted IS NULL 
            GROUP BY 1, 2, 3, 4
        """
        res_gps = client.query(q_gps).result()
        
        # Index by student_id and clean student_name
        grading_counts = {} # { "8": [ {station, body_part, cnt}, ... ], "張明暉": [...] }
        for r in res_gps:
            sid = str(r.sid).split('.')[0]
            sname = str(r.sname).strip().replace(' ', '').replace('　', '')
            entry = {"station": r.station, "body_part": r.body_part, "cnt": r.cnt}
            
            if sid:
                if sid not in grading_counts: grading_counts[sid] = []
                grading_counts[sid].append(entry)
            if sname:
                if sname not in grading_counts: grading_counts[sname] = []
                grading_counts[sname].append(entry)
            
        # 2. Aggregated Feedback Logs: Count per (student_name, department)
        q_fps = f"""
            SELECT student_name as sname, department as dept, COUNT(*) as cnt 
            FROM `{project}.grading_data.feedback_logs` 
            WHERE is_deleted = FALSE OR is_deleted IS NULL 
            GROUP BY 1, 2
        """
        res_fps = client.query(q_fps).result()
        feedback_counts = {}
        for r in res_fps:
            sname = str(r.sname).strip().replace(' ', '').replace('　', '')
            if sname:
                if sname not in feedback_counts: feedback_counts[sname] = []
                feedback_counts[sname].append({"dept": r.dept, "cnt": r.cnt})
            
        return grading_counts, feedback_counts
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {}, {}

def get_student_gamification_data(gc, doc, student_info):
    eps = doc.worksheet('各類別EPA需求').get_all_values()
    rps = doc.worksheet('檢查室清單').get_all_values()
    grading_counts, feedback_counts = get_bq_gamification_logs()
    exps = parse_exemptions(doc)
    scoring_config = parse_scoring_config(doc)
    
    s_email = ""
    st_vals = doc.worksheet('學員名單').get_all_values()
    for r in st_vals[1:]:
        if len(r) > 2 and str(r[0]).strip() == str(student_info.get('id', '')).strip():
            s_email = str(r[2]).strip()
            break
            
    # Standardize data for process function
    indexed_data = {'grades': grading_counts, 'fb': feedback_counts}
    
    return process_student_gamification(
        student_info.get('id', ''), student_info.get('name', ''), student_info.get('type', ''), student_info.get('gender', ''), s_email,
        eps, [], rps, [], [], exps, get_first_monday(), datetime.date.today(), scoring_config, indexed_data
    )

def get_leaderboard_data(gc, doc):
    try:
        eps = doc.worksheet('各類別EPA需求').get_all_values()
        rps = doc.worksheet('檢查室清單').get_all_values()
        exps = parse_exemptions(doc)
        scoring_config = parse_scoring_config(doc)
        fm = get_first_monday()
        today = datetime.date.today()
        
        st_vals = doc.worksheet('學員名單').get_all_values()
        header = st_vals[0]
        id_idx = header.index('學生ID') if '學生ID' in header else 0
        name_idx = header.index('姓名') if '姓名' in header else 1
        role_idx = header.index('學員類別') if '學員類別' in header else 4
        
        grading_counts, feedback_counts = get_bq_gamification_logs()
        indexed_data = {'grades': grading_counts, 'fb': feedback_counts}
        
        leaderboard = []
        for r in st_vals[1:]:
            v = {
                'id': str(r[id_idx]).strip(),
                'name': str(r[name_idx]).strip(),
                'type': str(r[role_idx]).strip() if len(r) > role_idx else '',
                'gender': str(r[header.index('性別')]).strip() if '性別' in header else '',
                'email': str(r[2]).strip() if len(r) > 2 else ''
            }
            # Only include person if name exists
            if v['name']:
                # Pass the indexed data to avoid scanning the logs for each student
                # We pass empty lists for logs because indexed_data is now our primary source
                data = process_student_gamification(v['id'], v['name'], v['type'], v['gender'], v['email'],
                                                   eps, [], rps, [], [], exps, fm, today, scoring_config, indexed_data)
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
