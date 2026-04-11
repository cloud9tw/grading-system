from google.cloud import bigquery
from google.oauth2 import service_account

def main():
    creds = service_account.Credentials.from_service_account_file('credentials.json')
    client = bigquery.Client(credentials=creds, project=creds.project_id)
    
    for tn in ['grading_logs', 'feedback_logs']:
        try:
            table = client.get_table(f'{creds.project_id}.grading_data.{tn}')
            print(f"--- {tn} ---")
            for field in table.schema:
                print(f"{field.name}: {field.field_type}")
        except Exception as e:
            print(f"Error {tn}: {e}")

if __name__ == "__main__":
    main()
