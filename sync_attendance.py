import os
import gspread
import datetime
from google.cloud import bigquery
from google.oauth2 import service_account
from dotenv import load_dotenv

def sync():
    print("Starting Attendance Sync...")
    load_dotenv()
    
    # 1. Auth
    creds_file = 'credentials.json'
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_file(creds_file, scopes=scopes)
    
    # Sheets
    gc = gspread.authorize(creds)
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    doc = gc.open_by_key(sheet_id)
    ws = doc.worksheet('上下班打卡記錄')
    vals = ws.get_all_values()
    
    if len(vals) < 2:
        print("No data in worksheet.")
        return

    header = vals[0]
    rows = vals[1:]
    
    # 2. BigQuery setup
    bq_creds = service_account.Credentials.from_service_account_file(creds_file)
    client = bigquery.Client(credentials=bq_creds, project=bq_creds.project_id)
    # We must write to the BASE table, not the view
    table_id = f"{bq_creds.project_id}.grading_data.attendance_events"
    
    # 3. Transform
    json_rows = []
    
    def parse_dt(ts_str):
        if not ts_str: return None
        for fmt in ["%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"]:
            try:
                dt = datetime.datetime.strptime(ts_str, fmt)
                return dt.isoformat()
            except: continue
        return None

    for r in rows:
        if not r or not r[0]: continue
        
        # Mapping: ['學生', '教師', '共同教師', '檢查室', '簽到時間', '簽退時間']
        base_data = {
            "student_name": r[0].strip(),
            "teacher_name": r[1].strip() if len(r) > 1 else '',
            "co_teacher": r[2].strip() if len(r) > 2 else '',
            "sub_room": r[3].strip() if len(r) > 3 else '',
            "is_deleted": False
        }
        
        in_time_str = r[4].strip() if len(r) > 4 else ''
        out_time_str = r[5].strip() if len(r) > 5 else ''
        
        # Create CHECK_IN event
        in_ts = parse_dt(in_time_str)
        if in_ts:
            ev_in = base_data.copy()
            ev_in.update({"event_type": "CHECK_IN", "event_time": in_ts})
            json_rows.append(ev_in)
            
        # Create CHECK_OUT event
        out_ts = parse_dt(out_time_str)
        if out_ts:
            ev_out = base_data.copy()
            ev_out.update({"event_type": "CHECK_OUT", "event_time": out_ts})
            json_rows.append(ev_out)

    print(f"Prepared {len(json_rows)} event rows for BigQuery.")

    # 4. Upload (Truncate & Load)
    job_config = bigquery.LoadJobConfig(
        write_disposition="WRITE_TRUNCATE",
        source_format=bigquery.SourceFormat.NEWLINE_DELIMITED_JSON,
    )
    
    try:
        job = client.load_table_from_json(json_rows, table_id, job_config=job_config)
        job.result()
        print(f"Successfully synced {len(json_rows)} events to {table_id}")
    except Exception as e:
        print(f"BQ Sync Error: {e}")

if __name__ == "__main__":
    sync()
