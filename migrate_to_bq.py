import os
import datetime
import re
from dotenv import load_dotenv
import gspread
from google.cloud import bigquery
from google.oauth2 import service_account

def migrate():
    print("Loading dotenv and setup...")
    load_dotenv()
    gc = gspread.service_account(filename='credentials.json')
    doc = gc.open_by_key(os.getenv("GOOGLE_SHEET_ID"))
    
    try:
        fb_doc = gc.open_by_key('112l_e3WKbIkFYj58nv8LRTYEvfyDpXMh-NcSe98T07w')
    except Exception as e:
        print("Error opening feedback doc:", e)
        return

    # BigQuery init
    print("Initializing BigQuery Client...")
    credentials = service_account.Credentials.from_service_account_file('credentials.json')
    bq_client = bigquery.Client(credentials=credentials, project=credentials.project_id)
    dataset_id = f"{credentials.project_id}.grading_data"

    # Truncate tables for fresh start
    for t in ['grading_logs', 'feedback_logs', 'attendance_events']:
        print(f"Truncating {t}...")
        try:
            bq_client.query(f"TRUNCATE TABLE `{dataset_id}.{t}`").result()
        except: pass

    def clean_id(val):
        if not val: return ""
        v = str(val).strip()
        if v.endswith('.0'): v = v[:-2]
        return v

    def parse_time(t_str):
        if not t_str: return None
        s = str(t_str).strip()
        
        # Handle Chinese AM/PM
        is_pm = False
        if '下午' in s: is_pm = True; s = s.replace('下午', ' ')
        elif '上午' in s: s = s.replace('上午', ' ')
        
        s = s.replace('/', '-')
        
        # Try many formats, including single digit month/day
        fmts = [
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
            "%Y-%m-%d %H:%M:%S", # redundant but safe
            "%Y-%n-%j", # not real specifiers but thinking about 2025-7-3
        ]
        
        # Best way for messy dates: split and reassemble
        try:
            # 2025-7-3 or 2025-7-3 11:22:10
            parts = s.split()
            date_parts = [int(x) for x in parts[0].split('-')]
            h, m, sec = 0, 0, 0
            if len(parts) > 1:
                time_parts = [int(x) for x in parts[1].split(':')]
                h = time_parts[0]
                m = time_parts[1]
                if len(time_parts) > 2: sec = time_parts[2]
            
            if is_pm and h < 12: h += 12
            if not is_pm and h == 12: h = 0
            
            dt = datetime.datetime(date_parts[0], date_parts[1], date_parts[2], h, m, sec)
            return dt.isoformat()
        except Exception as e:
            # print(f"Parse Error for [{t_str}]: {e}")
            return None

    # 1. Attendance
    print("Migrating Attendance...")
    att_vals = doc.worksheet('上下班打卡記錄').get_all_values()
    att_insert = []
    for r in att_vals[1:]:
        if len(r) >= 5:
            s_name, t_name, c_name, room, cin = r[0:5]
            cout = r[5] if len(r) > 5 else ''
            cin_iso = parse_time(cin)
            if cin_iso:
                att_insert.append({"student_name": s_name, "teacher_name": t_name, "co_teacher": c_name, "sub_room": room, "event_type": "CHECK_IN", "event_time": cin_iso, "is_deleted": False})
            cout_iso = parse_time(cout)
            if cout_iso:
                att_insert.append({"student_name": s_name, "teacher_name": t_name, "co_teacher": c_name, "sub_room": room, "event_type": "CHECK_OUT", "event_time": cout_iso, "is_deleted": False})
    
    if att_insert:
        bq_client.insert_rows_json(f"{dataset_id}.attendance_events", att_insert)
        print(f"Migrated {len(att_insert)} attendance events.")

    # 2. Grading
    print("Migrating Grading...")
    grade_vals = doc.worksheet('評分記錄').get_all_values()
    grade_insert = []
    for r in grade_vals[1:]:
        if len(r) > 6:
            ts_iso = parse_time(r[4])
            if ts_iso:
                grade_insert.append({
                    "timestamp": ts_iso,
                    "student_id": clean_id(r[0]),
                    "student_name": str(r[1]),
                    "station": str(r[2]),
                    "body_part": str(r[3]),
                    "opa1_sum": str(r[6]),
                    "opa2_sum": str(r[7]),
                    "opa3_sum": str(r[8]),
                    "opa1_items": [str(x) for x in r[9:17]],
                    "opa2_items": [str(x) for x in r[17:25]],
                    "opa3_items": [str(x) for x in r[25:33]],
                    "comment": str(r[35]) if len(r) > 35 else "",
                    "aspect1": str(r[33]) if len(r) > 33 else "",
                    "aspect2": str(r[34]) if len(r) > 34 else "",
                    "teacher_name": str(r[5]),
                    "is_deleted": False
                })
    if grade_insert:
        bq_client.insert_rows_json(f"{dataset_id}.grading_logs", grade_insert)
        print(f"Migrated {len(grade_insert)} grading logs.")

    # 3. Feedback
    print("Migrating Feedback...")
    fb_vals = fb_doc.worksheet('表單回應').get_all_values()
    fb_insert = []
    for r in fb_vals[1:]:
        if len(r) > 1:
            ts_iso = parse_time(r[1])
            if ts_iso:
                fb_insert.append({
                    "timestamp": ts_iso,
                    "email": str(r[3]),
                    "student_name": str(r[2]),
                    "role": "實習學生",
                    "teacher": str(r[4]),
                    "co_teacher": str(r[5]),
                    "department": str(r[6]),
                    "is_retake": "FALSE",
                    "score": str(r[32]) if len(r) > 32 else "",
                    "suggestions": str(r[33]) if len(r) > 33 else "",
                    "is_deleted": False
                })
    if fb_insert:
        bq_client.insert_rows_json(f"{dataset_id}.feedback_logs", fb_insert)
        print(f"Migrated {len(fb_insert)} feedback logs.")

    print("Migration Complete!")

if __name__ == "__main__":
    migrate()
