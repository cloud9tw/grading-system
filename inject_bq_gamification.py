import os

with open('gamification.py', 'r', encoding='utf-8') as f:
    content = f.read()

new_header = '''
def get_bq_gamification_logs():
    from google.cloud import bigquery
    from google.oauth2 import service_account
    from dotenv import load_dotenv
    import os
    
    load_dotenv()
    try:
        credentials = service_account.Credentials.from_service_account_file('credentials.json')
        bq = bigquery.Client(credentials=credentials, project=credentials.project_id)
        project = credentials.project_id
        
        q_gps = f\"\"\"
        SELECT student_id, student_name, station, body_part, timestamp, teacher_name, opa1_sum, opa2_sum, opa3_sum,
        opa1_items, opa2_items, opa3_items, aspect1, aspect2, comment
        FROM \{project}.grading_data.grading_logs\
        WHERE is_deleted = FALSE OR is_deleted IS NULL
        \"\"\"
        res = bq.query(q_gps).result()
        gps = [['學員ID', '學員姓名', '站別', '檢查部位', '時間', '教師姓名', 'OPA1總評', 'OPA2總評', 'OPA3總評', 'OPA1_1', 'OPA1_2', '...', 'OPA3_8', '面向選擇1', '面向選擇2', '簡評']]
        for r in res:
            row = [r.student_id, r.student_name, r.station, r.body_part, r.timestamp.strftime('%Y/%m/%d %H:%M:%S') if r.timestamp else '', r.teacher_name, r.opa1_sum, r.opa2_sum, r.opa3_sum]
            row.extend(list(r.opa1_items) if r.opa1_items else ['']*8)
            row.extend(list(r.opa2_items) if r.opa2_items else ['']*8)
            row.extend(list(r.opa3_items) if r.opa3_items else ['']*8)
            row.extend([r.aspect1, r.aspect2, r.comment])
            gps.append(row)

        q_fps = f\"\"\"
        SELECT timestamp, email, student_name, role, teacher, co_teacher, department, is_retake, score, suggestions
        FROM \{project}.grading_data.feedback_logs\
        WHERE is_deleted = FALSE OR is_deleted IS NULL
        \"\"\"
        res = bq.query(q_fps).result()
        fps = [['序號', '時間戳記', '學生姓名', '電子郵件地址', '教師名稱', '未登錄之教師姓名', '臨床實習站別', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '對教師的整體建議', '未登錄之教師姓名']]
        for r in res:
            row = ['', r.timestamp.strftime('%Y-%m-%d %H:%M:%S') if r.timestamp else '', r.student_name, r.email, r.teacher, r.co_teacher, r.department]
            row.extend(['']*26)
            row.extend([r.suggestions, r.co_teacher])
            fps.append(row)

        q_aps = f\"\"\"
        SELECT student_name, teacher_name, co_teacher, sub_room, check_in_time, check_out_time
        FROM \{project}.grading_data.attendance_daily_summary\
        \"\"\"
        res = bq.query(q_aps).result()
        aps = [['學生', '教師', '共同教師', '檢查室', '簽到時間', '簽退時間']]
        for r in res:
            aps.append([r.student_name, r.teacher_name, r.co_teacher, r.sub_room, r.check_in_time.strftime('%Y-%m-%d %H:%M:%S') if r.check_in_time else '', r.check_out_time.strftime('%Y-%m-%d %H:%M:%S') if r.check_out_time else ''])
            
        return gps, fps, aps
    except Exception as e:
        print('BQ Fetch Error:', e)
        return [], [], []

def get_student_gamification_data(gc, doc, student_info):
'''

old_header = '''
def get_student_gamification_data(gc, doc, student_info):
'''

content = content.replace(old_header.strip(), new_header.strip())

replace_logic = r'''    eps = doc.worksheet('各類別EPA需求').get_all_values()
    gps = doc.worksheet('評分記錄').get_all_values()
    rps = doc.worksheet('檢查室清單').get_all_values()
    import os
    fb_doc = gc.open_by_key('112l_e3WKbIkFYj58nv8LRTYEvfyDpXMh-NcSe98T07w')
    fps = fb_doc.worksheet('表單回應').get_all_values()
    aps = doc.worksheet('上下班打卡記錄').get_all_values()'''

new_logic = r'''    eps = doc.worksheet('各類別EPA需求').get_all_values()
    rps = doc.worksheet('檢查室清單').get_all_values()
    gps, fps, aps = get_bq_gamification_logs()'''

content = content.replace(replace_logic, new_logic)

with open('gamification.py', 'w', encoding='utf-8') as f:
    f.write(content)
print('Done injecting gamification')
