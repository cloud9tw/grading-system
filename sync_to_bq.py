# sync_to_bq.py
# 2026/04/22 Update: Modularized for BQ-centric architecture.

import os
import gspread
import datetime
from google.cloud import bigquery
from google.oauth2 import service_account
from dotenv import load_dotenv
from privacy_utils import get_code

def parse_dt(ts_str):
    if not ts_str: return None
    # Support various formats: YYYY/MM/DD, YYYY-MM-DD, with or without time
    formats = [
        "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M",
        "%Y-%m-%d", "%Y/%m/%d"
    ]
    for fmt in formats:
        try:
            dt = datetime.datetime.strptime(ts_str.strip(), fmt)
            return dt.isoformat()
        except: continue
    return None

def sync_all(callback=None):
    """
    Synchronizes Grading Logs and Attendance from Sheets to BigQuery.
    Anonymizes names and IDs before upload.
    """
    def log(msg):
        print(msg)
        if callback:
            # Check if callback is async
            import inspect
            if inspect.iscoroutinefunction(callback):
                pass # Caller should handle awaitable loggers if needed
            else:
                callback(msg)

    log("Starting Anonymized Sync Process: [Sheets -> BigQuery]")
    load_dotenv()
    
    # Auth setup
    base_path = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
    creds_file = os.path.join(base_path, 'credentials.json')
    
    if not os.path.exists(creds_file):
        log(f"Error: Credentials file not found at {creds_file}")
        return False

    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = service_account.Credentials.from_service_account_file(creds_file, scopes=scopes)
        gc = gspread.authorize(creds)
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        doc = gc.open_by_key(sheet_id)
        
        bq_creds = service_account.Credentials.from_service_account_file(creds_file)
        client = bigquery.Client(credentials=bq_creds, project=bq_creds.project_id)
        project = bq_creds.project_id

        # --- 1. SYNC ATTENDANCE ---
        log("[1/2] Anonymizing & Syncing Attendance Records...")
        try:
            ws = doc.worksheet('上下班打卡記錄')
            vals = ws.get_all_values()
            if len(vals) > 1:
                rows = vals[1:]
                json_rows = []
                for r in rows:
                    if not r or not r[0]: continue
                    
                    # 去識別化處理
                    s_name = r[0].strip()
                    t_name = r[1].strip() if len(r) > 1 else ''
                    
                    anon_s_code = get_code(s_name, 'student')
                    anon_t_code = get_code(t_name, 'teacher') if t_name else ''

                    base = {
                        "student_name": anon_s_code,
                        "teacher_name": anon_t_code,
                        "co_teacher": r[2].strip() if len(r) > 2 else '',
                        "sub_room": r[3].strip() if len(r) > 3 else '',
                        "is_deleted": False
                    }
                    in_ts = parse_dt(r[4] if len(r) > 4 else '')
                    if in_ts:
                        json_rows.append({**base, "event_type": "CHECK_IN", "event_time": in_ts})
                    out_ts = parse_dt(r[5] if len(r) > 5 else '')
                    if out_ts:
                        json_rows.append({**base, "event_type": "CHECK_OUT", "event_time": out_ts})
                
                if json_rows:
                    job_config = bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE", source_format="NEWLINE_DELIMITED_JSON")
                    table_id = f"{project}.grading_data.attendance_events"
                    client.load_table_from_json(json_rows, table_id, job_config=job_config).result()
                    log(f"Success: Attendance Sync ({len(json_rows)} events, ANONYMIZED)")
            else:
                log("Attendance Sync: No data found.")
        except Exception as e:
            log(f"Attendance Sync Failed: {str(e)}")

        # --- 2. SYNC GRADING LOGS ---
        log("[2/2] Anonymizing & Syncing EPA Grading Logs...")
        try:
            ws = doc.worksheet('評分記錄')
            vals = ws.get_all_values()
            if len(vals) > 1:
                rows = vals[1:]
                json_rows = []
                for r in rows:
                    if not r or not (len(r) > 4 and r[4]): continue
                    ts = parse_dt(r[4].strip())
                    if not ts: continue
                    
                    def join_opa(start_idx):
                        if len(r) <= start_idx: return ""
                        items = r[start_idx : start_idx+8]
                        return ",".join([str(x) for x in items if x])

                    # 去識別化處理
                    raw_sname = r[1].strip()
                    raw_tname = r[5].strip() if len(r) > 5 else ''
                    
                    anon_s_code = get_code(raw_sname, 'student')
                    anon_t_code = get_code(raw_tname, 'teacher') if raw_tname else ''

                    json_rows.append({
                        "student_id": anon_s_code,
                        "student_name": anon_s_code,
                        "station": r[2].strip(),
                        "body_part": r[3].strip(),
                        "timestamp": ts,
                        "teacher_name": anon_t_code,
                        "opa1_sum": r[6].strip() if len(r) > 6 else '',
                        "opa2_sum": r[7].strip() if len(r) > 7 else '',
                        "opa3_sum": r[8].strip() if len(r) > 8 else '',
                        "opa1_items": join_opa(9),
                        "opa2_items": join_opa(17),
                        "opa3_items": join_opa(25),
                        "aspect1": r[33].strip() if len(r) > 33 else '',
                        "aspect2": r[34].strip() if len(r) > 34 else '',
                        "comment": r[35].strip() if len(r) > 35 else '',
                        "is_deleted": False
                    })
                
                if json_rows:
                    job_config = bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE", source_format="NEWLINE_DELIMITED_JSON")
                    table_id = f"{project}.grading_data.grading_logs"
                    client.load_table_from_json(json_rows, table_id, job_config=job_config).result()
                    log(f"Success: EPA Logs Sync ({len(json_rows)} entries, ANONYMIZED)")
            else:
                log("EPA Logs Sync: No data found.")
        except Exception as e:
            log(f"EPA Logs Sync Failed: {str(e)}")

        # --- 3. SYNC CEEP DOPS ---
        log("[3/4] Syncing CEEP DOPS Records...")
        try:
            ws = doc.worksheet('CEEP_DOPS')
            vals = ws.get_all_values()
            if len(vals) > 1:
                json_rows = []
                for r in vals[1:]:
                    if len(r) < 27: continue
                    ts = parse_dt(r[3]) # 假設索引 3 是提交時間
                    json_rows.append({
                        "timestamp": ts,
                        "student_name": r[2].strip(),
                        "station": r[11].strip(),
                        "score": float(r[26]) if r[26] else 0.0,
                        "feedback": r[23].strip(),
                        "raw_data": ""
                    })
                if json_rows:
                    job_config = bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE", source_format="NEWLINE_DELIMITED_JSON")
                    table_id = f"{project}.grading_data.dops_logs"
                    client.load_table_from_json(json_rows, table_id, job_config=job_config).result()
                    log(f"Success: CEEP DOPS Sync ({len(json_rows)} entries)")
        except Exception as e:
            log(f"CEEP DOPS Sync Failed: {str(e)}")

        # --- 4. SYNC CEEP MiniCEX ---
        log("[4/4] Syncing CEEP MiniCEX Records...")
        try:
            ws = doc.worksheet('CEEP_MiniCEX')
            vals = ws.get_all_values()
            if len(vals) > 1:
                json_rows = []
                for r in vals[1:]:
                    if len(r) < 22: continue
                    ts = parse_dt(r[3])
                    json_rows.append({
                        "timestamp": ts,
                        "student_name": r[2].strip(),
                        "station": r[10].strip(), # 這裡先用原始欄位，由後端 clean
                        "score": float(r[21]) if r[21] else 0.0,
                        "feedback": r[20].strip()
                    })
                if json_rows:
                    job_config = bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE", source_format="NEWLINE_DELIMITED_JSON")
                    table_id = f"{project}.grading_data.minicex_logs"
                    client.load_table_from_json(json_rows, table_id, job_config=job_config).result()
                    log(f"Success: CEEP MiniCEX Sync ({len(json_rows)} entries)")
        except Exception as e:
            log(f"CEEP MiniCEX Sync Failed: {str(e)}")

        log("Sync Complete: BigQuery is now Up-to-Date.")
        return True
    except Exception as e:
        log(f"Full Sync FATAL Error: {str(e)}")
        return False

if __name__ == "__main__":
    sync_all()
