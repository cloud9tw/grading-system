
import os
from google.cloud import bigquery
from google.oauth2 import service_account
from dotenv import load_dotenv

def check_bq():
    base_path = r'c:\Users\cloud\Desktop\EPA-grading\grading-system'
    load_dotenv(os.path.join(base_path, '.env'))
    creds_file = os.path.join(base_path, 'credentials.json')
    
    creds = service_account.Credentials.from_service_account_file(creds_file)
    client = bigquery.Client(credentials=creds, project=creds.project_id)
    project_id = creds.project_id
    
    target_date = "2026-04-12"
    target_student = "盧仁偉"
    target_station = "CT"
    
    q = f"""
        SELECT timestamp, station, student_name, opa1_sum, comment
        FROM `{project_id}.grading_data.grading_logs`
        WHERE student_name = '{target_student}' AND station = '{target_station}'
        ORDER BY timestamp ASC
    """
    query_job = client.query(q)
    results = query_job.result()
    
    found = False
    print(f"--- Records for {target_student} in {target_station} ---")
    for r in results:
        found = True
        print(f"Time: {r.timestamp} | Station: {r.station} | Comment: {r.comment}")
        
    if not found:
        print("No records found in BigQuery.")

if __name__ == "__main__":
    check_bq()
