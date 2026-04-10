import os
import json
import datetime
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
    
    # 1. Dataset
    try:
        bq_client.get_dataset(dataset_id)
        print(f"Dataset {dataset_id} already exists.")
    except Exception:
        dataset = bigquery.Dataset(dataset_id)
        dataset.location = "asia-east1"
        bq_client.create_dataset(dataset, timeout=30)
        print(f"Created dataset {dataset_id}")

    # 2. Schema Definitions
    att_schema = [
        bigquery.SchemaField("student_name", "STRING"),
        bigquery.SchemaField("teacher_name", "STRING"),
        bigquery.SchemaField("co_teacher", "STRING"),
        bigquery.SchemaField("sub_room", "STRING"),
        bigquery.SchemaField("event_type", "STRING"),
        bigquery.SchemaField("event_time", "TIMESTAMP"),
    ]
    
    grade_schema = [
        bigquery.SchemaField("timestamp", "TIMESTAMP"),
        bigquery.SchemaField("student_id", "STRING"),
        bigquery.SchemaField("station", "STRING"),
        bigquery.SchemaField("body_part", "STRING"),
        bigquery.SchemaField("opa1_sum", "STRING"),
        bigquery.SchemaField("opa2_sum", "STRING"),
        bigquery.SchemaField("opa3_sum", "STRING"),
        bigquery.SchemaField("opa1_items", "STRING", mode="REPEATED"),
        bigquery.SchemaField("opa2_items", "STRING", mode="REPEATED"),
        bigquery.SchemaField("opa3_items", "STRING", mode="REPEATED"),
        bigquery.SchemaField("comment", "STRING"),
        bigquery.SchemaField("teacher_name", "STRING"),
        bigquery.SchemaField("status", "STRING")
    ]
    
    fb_schema = [
        bigquery.SchemaField("timestamp", "TIMESTAMP"),
        bigquery.SchemaField("email", "STRING"),
        bigquery.SchemaField("student_name", "STRING"),
        bigquery.SchemaField("role", "STRING"),
        bigquery.SchemaField("teacher", "STRING"),
        bigquery.SchemaField("co_teacher", "STRING"),
        bigquery.SchemaField("department", "STRING"),
        bigquery.SchemaField("is_retake", "STRING"),
        bigquery.SchemaField("score", "STRING"),
        bigquery.SchemaField("suggestions", "STRING"),
    ]

    # --- Table Creations ---
    att_table_id = f"{dataset_id}.attendance_events"
    try:
        bq_client.get_table(att_table_id)
    except Exception:
        table = bigquery.Table(att_table_id, schema=att_schema)
        bq_client.create_table(table)
        print(f"Created table {att_table_id}")

    grade_table_id = f"{dataset_id}.grading_logs"
    try:
        bq_client.get_table(grade_table_id)
    except Exception:
        table = bigquery.Table(grade_table_id, schema=grade_schema)
        table.time_partitioning = bigquery.TimePartitioning(type_=bigquery.TimePartitioningType.DAY, field="timestamp")
        table.clustering_fields = ["teacher_name", "station"]
        bq_client.create_table(table)
        print(f"Created table {grade_table_id}")

    fb_table_id = f"{dataset_id}.feedback_logs"
    try:
        bq_client.get_table(fb_table_id)
    except Exception:
        table = bigquery.Table(fb_table_id, schema=fb_schema)
        table.time_partitioning = bigquery.TimePartitioning(type_=bigquery.TimePartitioningType.DAY, field="timestamp")
        table.clustering_fields = ["teacher", "department"]
        bq_client.create_table(table)
        print(f"Created table {fb_table_id}")

    # --- Data Migration ---
    # Convert 'yyyy/mm/dd' Google sheet datetime strings to ISO format valid for BigQuery
    def parse_time(t_str):
        if not t_str: return None
        t_str = t_str.strip()
        # Usually '%Y-%m-%d %H:%M:%S' or '%Y/%m/%d %H:%M:%S'
        t_str = t_str.replace('/', '-')
        try:
            return datetime.datetime.strptime(t_str, "%Y-%m-%d %H:%M:%S").isoformat()
        except:
            return None

    print("Migrating Attendance...")
    att_vals = doc.worksheet('上下班打卡記錄').get_all_values()
    att_insert = []
    for r in att_vals[1:]:
        if len(r) >= 6:
            s_name, t_name, c_name, room, cin, cout = r[0:6]
            cin_iso = parse_time(cin)
            if cin_iso:
                att_insert.append({"student_name": s_name, "teacher_name": t_name, "co_teacher": c_name, "sub_room": room, "event_type": "CHECK_IN", "event_time": cin_iso})
            cout_iso = parse_time(cout)
            if cout_iso:
                att_insert.append({"student_name": s_name, "teacher_name": t_name, "co_teacher": c_name, "sub_room": room, "event_type": "CHECK_OUT", "event_time": cout_iso})
                
    if att_insert:
        errors = bq_client.insert_rows_json(att_table_id, att_insert)
        if errors: print("Attendance inserting errors:", errors)
        else: print(f"Migrated {len(att_insert)} attendance event records.")
    
    print("Migrating Grading...")
    grade_vals = doc.worksheet('評分記錄').get_all_values()
    grade_insert = []
    for r in grade_vals[1:]:
        # row expects 34 cols
        # pad if shorter
        r += [''] * (34 - len(r))
        ts_iso = parse_time(r[0])
        if ts_iso:
            grade_insert.append({
                "timestamp": ts_iso,
                "student_id": r[1],
                "station": r[2],
                "body_part": r[3],
                "opa1_sum": r[4],
                "opa2_sum": r[5],
                "opa3_sum": r[6],
                "opa1_items": r[7:15],
                "opa2_items": r[15:23],
                "opa3_items": r[23:31],
                "comment": r[31],
                "teacher_name": r[32] if r[32] else r[1], # some fallback
                "status": r[33]
            })
    if grade_insert:
        errors = bq_client.insert_rows_json(grade_table_id, grade_insert)
        if errors: print("Grading inserting errors:", errors)
        else: print(f"Migrated {len(grade_insert)} grading records.")

    print("Migrating Feedback...")
    fb_vals = fb_doc.worksheet('表單回應').get_all_values()
    fb_insert = []
    for r in fb_vals[1:]:
        r += [''] * (10 - len(r))
        ts_iso = parse_time(r[0])
        if ts_iso:
            fb_insert.append({
                "timestamp": ts_iso,
                "email": r[1],
                "student_name": r[2],
                "role": r[3],
                "teacher": r[4],
                "co_teacher": r[5],
                "department": r[6],
                "is_retake": r[7],
                "score": r[8],
                "suggestions": r[9]
            })
    if fb_insert:
        errors = bq_client.insert_rows_json(fb_table_id, fb_insert)
        if errors: print("Feedback inserting errors:", errors)
        else: print(f"Migrated {len(fb_insert)} feedback records.")

    print("Creating View: attendance_daily_summary")
    view_id = f"{dataset_id}.attendance_daily_summary"
    view_sql = f"""
    SELECT 
      student_name,
      teacher_name,
      co_teacher,
      sub_room,
      DATE(event_time) AS event_date,
      MIN(CASE WHEN event_type = 'CHECK_IN' THEN event_time END) AS check_in_time,
      MAX(CASE WHEN event_type = 'CHECK_OUT' THEN event_time END) AS check_out_time
    FROM `{dataset_id}.attendance_events`
    GROUP BY student_name, teacher_name, co_teacher, sub_room, DATE(event_time)
    ORDER BY event_date DESC
    """
    try:
        view = bigquery.Table(view_id)
        view.view_query = view_sql
        bq_client.create_table(view, exists_ok=True)
        print("View created or updated.")
    except Exception as e:
        print("Error creating view:", e)

    print("Migration complete!")

if __name__ == "__main__":
    migrate()
