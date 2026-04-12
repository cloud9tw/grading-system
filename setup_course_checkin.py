import os
import gspread
from google.cloud import bigquery
from google.oauth2 import service_account
from dotenv import load_dotenv

def setup_infra():
    load_dotenv()
    print("🚀 Initializing Course Check-in Infrastructure...")
    
    # 1. Update Google Sheet Structure
    try:
        creds_file = 'credentials.json'
        scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = service_account.Credentials.from_service_account_file(creds_file, scopes=scopes)
        gc = gspread.authorize(creds)
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        doc = gc.open_by_key(sheet_id)
        ws = doc.worksheet('早8課程簽到(含放腫、全人教學)')
        
        headers = ws.row_values(1)
        if '上課日期' not in headers:
            print("  - Inserting Date and Time columns into Google Sheet...")
            # Insert at column 3 (after Teacher)
            ws.insert_cols([['上課日期'], ['開始時間']], 3)
            print("  ✅ Sheet structure updated.")
        else:
            print("  - Sheet structure already correct.")
    except Exception as e:
        print(f"  ❌ Sheet setup failed: {e}")

    # 2. Initialize BigQuery Table
    try:
        bq_creds = service_account.Credentials.from_service_account_file(creds_file)
        client = bigquery.Client(credentials=bq_creds, project=bq_creds.project_id)
        dataset_id = f"{bq_creds.project_id}.grading_data"
        table_id = f"{dataset_id}.course_checkins"
        
        schema = [
            bigquery.SchemaField("student_id", "STRING", mode="REQUIRED"),
            bigquery.SchemaField("student_name", "STRING", mode="REQUIRED"),
            bigquery.SchemaField("course_name", "STRING", mode="REQUIRED"),
            bigquery.SchemaField("hours", "FLOAT", mode="REQUIRED"),
            bigquery.SchemaField("is_manual", "BOOLEAN", mode="REQUIRED"),
            bigquery.SchemaField("timestamp", "TIMESTAMP", mode="REQUIRED"),
        ]
        
        try:
            client.get_table(table_id)
            print(f"  - BigQuery table '{table_id}' already exists.")
        except:
            print(f"  - Creating BigQuery table '{table_id}'...")
            table = bigquery.Table(table_id, schema=schema)
            client.create_table(table)
            print("  ✅ BigQuery table created.")
            
    except Exception as e:
        print(f"  ❌ BigQuery setup failed: {e}")

if __name__ == "__main__":
    setup_infra()
