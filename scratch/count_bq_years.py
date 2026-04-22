
import os
from google.cloud import bigquery
from google.oauth2 import service_account
from dotenv import load_dotenv

def count_bq():
    base_path = r'c:\Users\cloud\Desktop\EPA-grading\grading-system'
    load_dotenv(os.path.join(base_path, '.env'))
    creds_file = os.path.join(base_path, 'credentials.json')
    creds = service_account.Credentials.from_service_account_file(creds_file)
    client = bigquery.Client(credentials=creds, project=creds.project_id)
    project_id = creds.project_id
    
    q = f"SELECT EXTRACT(YEAR FROM timestamp) as yr, COUNT(*) as cnt FROM `{project_id}.grading_data.grading_logs` GROUP BY yr"
    res = client.query(q).result()
    print("--- BigQuery Year Distribution ---")
    for r in res:
        print(f"Year: {r.yr} | Count: {r.cnt}")

if __name__ == "__main__":
    count_bq()
