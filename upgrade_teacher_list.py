import os
import gspread
from google.oauth2 import service_account
from dotenv import load_dotenv

def upgrade():
    load_dotenv()
    print("🚀 Upgrading Teacher List structure...")
    try:
        creds_file = 'credentials.json'
        scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = service_account.Credentials.from_service_account_file(creds_file, scopes=scopes)
        gc = gspread.authorize(creds)
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        doc = gc.open_by_key(sheet_id)
        ws = doc.worksheet('教師名單')
        
        headers = ws.row_values(1)
        if '管理員權限' not in headers:
            print("  - Adding '管理員權限' column to E1...")
            ws.update_cell(1, len(headers) + 1, '管理員權限')
            print("  ✅ Column added successfully.")
        else:
            print("  - Column '管理員權限' already exists.")
    except Exception as e:
        print(f"  ❌ Upgrade failed: {e}")

if __name__ == "__main__":
    upgrade()
