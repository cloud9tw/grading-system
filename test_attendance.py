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
        
        try:
            print("--- 檢查室清單 ---")
            st_sheet = doc.worksheet('檢查室清單')
            for row in st_sheet.get_all_values()[:5]:
                print(row)
        except Exception as e:
            print("檢查室清單 NOT FOUND", e)
            
        try:
            print("\n--- 上下班打卡記錄 ---")
            log_sheet = doc.worksheet('上下班打卡記錄')
            for row in log_sheet.get_all_values()[:2]:
                print(row)
        except Exception as e:
            print("上下班打卡記錄 NOT FOUND", e)

if __name__ == '__main__':
    main()
