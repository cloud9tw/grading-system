
import os
from google.cloud import bigquery
from google.oauth2 import service_account
from dotenv import load_dotenv

def check_2026():
    base_path = r'c:\Users\cloud\Desktop\EPA-grading\grading-system'
    load_dotenv(os.path.join(base_path, '.env'))
    creds_file = os.path.join(base_path, 'credentials.json')
    creds = service_account.Credentials.from_service_account_file(creds_file)
    client = bigquery.Client(credentials=creds, project=creds.project_id)
    project_id = creds.project_id
    
    q = f"SELECT student_id, student_name, station, timestamp FROM `{project_id}.grading_data.grading_logs` WHERE EXTRACT(YEAR FROM timestamp) = 2026"
    res = client.query(q).result()
    print("--- 2026 Records in BigQuery ---")
    for r in res:
        print(f"ID: {r.student_id} | Name: {r.student_name} | Station: {r.station} | Time: {r.timestamp}")

if __name__ == "__main__":
    check_2026()
