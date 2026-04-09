import datetime
import re

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

def process_student_gamification(s_name, s_type, s_gender, s_email, epa_vals, grade_vals, room_vals, fb_vals, attendance_vals, exemptions, first_monday, today):
    # === 1. EPA Progress ===
    epa_progress = {}
    type_idx = -1
    if epa_vals and len(epa_vals) > 0:
        for i, val in enumerate(epa_vals[0]):
            if val.strip() == s_type:
                type_idx = i; break
    curr_cat = "未分類"
    if type_idx != -1:
        for r in epa_vals[1:]:
            station = str(r[0]).strip()
            if not station: continue
            if station.upper().startswith('EPA'):
                curr_cat = station; continue
            m = re.search(r'\d+', str(r[type_idx]) if type_idx < len(r) else '')
            if m:
                if curr_cat not in epa_progress: epa_progress[curr_cat] = {}
                epa_progress[curr_cat][station] = {'target': int(m.group(0)), 'current': 0}
                
    for r in grade_vals[1:]:
        if len(r) > 3 and str(r[1]).strip() == s_name:
            station_dept = str(r[2]).strip()
            body_part = str(r[3]).strip()
            for cat, stations in epa_progress.items():
                for epa_k, prog in stations.items():
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
            if len(r) > 6 and str(r[2]).strip() == s_name:
                dept = str(r[6]).strip()
                if '急診' in dept:
                    for k in feedback_progress.keys():
                        if 'Routine' in k:
                            feedback_progress[k]['current'] += 1; break
                else:
                    for k in feedback_progress.keys():
                        if k in dept or dept in k:
                            feedback_progress[k]['current'] += 1; break

    # === 3. Points & EPA/FB Medals ===
    points = 0
    medals = []
    
    # EPA
    epa_master = True
    has_epa_targets = False
    for cat, stations in epa_progress.items():
        cat_done = True
        for k, v in stations.items():
            has_epa_targets = True
            points += v['current'] * 10
            if v['current'] < v['target']:
                cat_done = False
                epa_master = False
        if has_epa_targets:
            name_prefix = cat.split('-')[0]
            medals.append({'name': f"{name_prefix}專精", 'desc': f"完成 {cat} 所有需求", 'achieved': cat_done})
            
    if has_epa_targets:
        medals.append({'name': 'EPA 大師', 'desc': '所有身分設定之 EPA 分類全數解鎖', 'achieved': epa_master})
        
    # Feedback
    feedback_master = True
    has_fb_targets = False
    for k, v in feedback_progress.items():
        has_fb_targets = True
        points += v['current'] * 5
        fb_done = (v['current'] >= v['target'])
        if not fb_done: feedback_master = False
        name_prefix = k.split('(')[-1].replace(')', '') if '(' in k else k
        medals.append({'name': f"{name_prefix}回饋", 'desc': f"繳交 {k} 教學回饋表單", 'achieved': fb_done})
        
    if has_fb_targets:
        medals.append({'name': '回饋達人', 'desc': '所有目標站別之回饋表單全數達成', 'achieved': feedback_master})
        
    # === 4. Attendance Medals ===
    internship_week = 0
    missed_months = set()
    perfect_months = 0
    overall_perfect = True
    
    if s_type == '實習學生':
        delta = today - first_monday
        if delta.days >= 0:
            internship_week = min(28, (delta.days // 7) + 1)
            
        attended_dates = {}
        for r in attendance_vals[1:]:
            if len(r) > 5 and str(r[0]).strip() == s_name:
                cin, cout = str(r[4]).strip(), str(r[5]).strip()
                d1, d2 = None, None
                if cin:
                    try: d1 = datetime.datetime.strptime(cin, "%Y-%m-%d %H:%M:%S").date()
                    except: pass
                if cout:
                    try: d2 = datetime.datetime.strptime(cout, "%Y-%m-%d %H:%M:%S").date()
                    except: pass
                if d1:
                    if d1 not in attended_dates: attended_dates[d1] = {'in': False, 'out': False}
                    attended_dates[d1]['in'] = True
                if d2:
                    if d2 not in attended_dates: attended_dates[d2] = {'in': False, 'out': False}
                    attended_dates[d2]['out'] = True
                    
        check_date = first_monday
        while check_date < today and check_date < first_monday + datetime.timedelta(weeks=28): # Not strictly < today, since today could be missing logout
            if check_date.weekday() < 5:
                is_exempt = False
                if check_date in exemptions:
                    if not exemptions[check_date] or s_email in exemptions[check_date]:
                        is_exempt = True
                if not is_exempt:
                    day_log = attended_dates.get(check_date, {'in': False, 'out': False})
                    # Must have both check-in and check-out
                    if not day_log['in'] or not day_log['out']:
                        missed_months.add((check_date.year, check_date.month))
                        overall_perfect = False
            check_date += datetime.timedelta(days=1)
            
        months_status = []
        for i in range(7):
            ty = first_monday.year + (first_monday.month + i - 1) // 12
            tm = (first_monday.month + i - 1) % 12 + 1
            if datetime.date(ty, tm, 1) > today: # Future month
                months_status.append(False)
            elif (ty, tm) in missed_months:
                months_status.append(False)
            else:
                months_status.append(True)
                perfect_months += 1
                
        for i in range(7):
            tm = (first_monday.month + i - 1) % 12 + 1
            medals.append({
                'name': f"{tm}月全勤", 
                'desc': f"於 {tm} 月落實打卡且無任何缺勤紀錄", 
                'achieved': months_status[i]
            })
            
        medals.append({
            'name': '28週均全勤', 
            'desc': '實習期間每一個工作日皆完美打卡出勤', 
            'achieved': overall_perfect and (internship_week >= 28)
        })

    achieved_count = sum(1 for m in medals if m['achieved'])

    return {
        'student_type': s_type,
        'gender': s_gender,
        'internship_week': internship_week,
        'points': points,
        'medals': medals,
        'epa_progress': epa_progress,
        'feedback_progress': feedback_progress,
        'achieved_count': achieved_count
    }

def get_first_monday():
    now = datetime.datetime.now()
    year = now.year
    july_first = datetime.date(year, 7, 1)
    days_to_monday = (0 - july_first.weekday() + 7) % 7
    first_monday = july_first + datetime.timedelta(days=days_to_monday)
    if now.date() < first_monday:
        july_first = datetime.date(year - 1, 7, 1)
        days_to_monday = (0 - july_first.weekday() + 7) % 7
        first_monday = july_first + datetime.timedelta(days=days_to_monday)
    return first_monday

def get_student_gamification_data(gc, doc, student_info):
    eps = doc.worksheet('各類別EPA需求').get_all_values()
    gps = doc.worksheet('評分記錄').get_all_values()
    rps = doc.worksheet('檢查室清單').get_all_values()
    import os
    fb_doc = gc.open_by_key('112l_e3WKbIkFYj58nv8LRTYEvfyDpXMh-NcSe98T07w')
    fps = fb_doc.worksheet('表單回應').get_all_values()
    aps = doc.worksheet('上下班打卡記錄').get_all_values()
    exps = parse_exemptions(doc)
    
    # Needs email
    s_email = ""
    st_vals = doc.worksheet('學員名單').get_all_values()
    for r in st_vals[1:]:
        if len(r) > 2 and str(r[0]).strip() == student_info.get('id', ''):
            s_email = str(r[2]).strip()
            break
            
    return process_student_gamification(
        student_info.get('name', ''), student_info.get('type', ''), student_info.get('gender', ''), s_email,
        eps, gps, rps, fps, aps, exps, get_first_monday(), datetime.date.today()
    )

def get_leaderboard_data(gc, doc):
    try:
        eps = doc.worksheet('各類別EPA需求').get_all_values()
        gps = doc.worksheet('評分記錄').get_all_values()
        rps = doc.worksheet('檢查室清單').get_all_values()
        import os
        fb_doc = gc.open_by_key('112l_e3WKbIkFYj58nv8LRTYEvfyDpXMh-NcSe98T07w')
        fps = fb_doc.worksheet('表單回應').get_all_values()
        aps = doc.worksheet('上下班打卡記錄').get_all_values()
        exps = parse_exemptions(doc)
        fm = get_first_monday()
        today = datetime.date.today()
        
        interns = []
        st_vals = doc.worksheet('學員名單').get_all_values()
        
        header = st_vals[0]
        gender_idx = header.index('性別') if '性別' in header else -1
        
        for r in st_vals[1:]:
            if len(r) > 4 and str(r[4]).strip() == '實習學生':
                interns.append({
                    'name': str(r[1]).strip(),
                    'email': str(r[2]).strip() if len(r) > 2 else '',
                    'type': '實習學生',
                    'gender': str(r[gender_idx]).strip() if gender_idx != -1 and gender_idx < len(r) else ''
                })
                
        leaderboard = []
        for v in interns:
            data = process_student_gamification(v['name'], v['type'], v['gender'], v['email'],
                                               eps, gps, rps, fps, aps, exps, fm, today)
            leaderboard.append({
                'name': v['name'],
                'points': data['points'],
                'medals': data['achieved_count']
            })
            
        leaderboard.sort(key=lambda x: x['points'], reverse=True)
        return leaderboard
    except Exception as e:
        print("Leaderboard count error:", e)
        return []
