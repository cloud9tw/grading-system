
import os
import gspread
from google.oauth2 import service_account
from dotenv import load_dotenv

def check_last_rows():
    base_path = r'c:\Users\cloud\Desktop\EPA-grading\grading-system'
    load_dotenv(os.path.join(base_path, '.env'))
    creds_file = os.path.join(base_path, 'credentials.json')
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_file(creds_file, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    doc = gc.open_by_key(sheet_id)
    ws = doc.worksheet('評分記錄')
    
    vals = ws.get_all_values()
    if len(vals) > 10:
        rows = vals[-10:]
    else:
        rows = vals[1:]
        
    print("--- Last 10 rows of 評分記錄 ---")
    for r in rows:
        print(r[:7]) # Print first 7 columns: ID, Name, Station, Body, Time, Teacher, OPA1 Sum

if __name__ == "__main__":
    check_last_rows()
