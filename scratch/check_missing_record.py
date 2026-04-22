
import os
import gspread
from google.oauth2 import service_account
from dotenv import load_dotenv

def check_sheet():
    base_path = r'c:\Users\cloud\Desktop\EPA-grading\grading-system'
    load_dotenv(os.path.join(base_path, '.env'))
    creds_file = os.path.join(base_path, 'credentials.json')
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_file(creds_file, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    doc = gc.open_by_key(sheet_id)
    ws = doc.worksheet('評分記錄')
    records = ws.get_all_records()
    
    target_date = "2026/04/12"
    target_student = "盧仁偉"
    target_station = "CT"
    
    found = []
    for r in records:
        s_name = str(r.get('學員姓名', '')).strip()
        stn = str(r.get('站別', '')).strip()
        dt = str(r.get('時間', '')).strip()
        
        if s_name == target_student and stn == target_station:
            found.append(r)
            print(f"Found record: {dt} | {s_name} | {stn} | Comment: {r.get('簡易評語', 'N/A')}")

    if not found:
        print("No record found in Google Sheet for 盧仁偉/CT.")
    else:
        print(f"Total found in Sheet: {len(found)}")

if __name__ == "__main__":
    check_sheet()
