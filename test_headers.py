import os
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials

load_dotenv()

def main():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = os.getenv("GOOGLE_APPLICATION_CREDENTIALS") or os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON") or "credentials.json"
    
    if os.path.exists(creds_file):
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
        gc = gspread.authorize(creds)
        
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        print("--- 站別OPA細項 ---")
        st_sheet = doc.worksheet('站別OPA細項')
        print(st_sheet.row_values(1))
        for row in st_sheet.get_all_values()[1:3]:
            print("Row data:", row)
            
        print("\n--- 評分記錄 ---")
        log_sheet = doc.worksheet('評分記錄')
        print("First row values:", log_sheet.row_values(1))
        
if __name__ == '__main__':
    main()
