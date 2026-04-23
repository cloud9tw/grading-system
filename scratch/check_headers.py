
import os
import gspread
from google.oauth2 import service_account
from dotenv import load_dotenv

def get_headers():
    load_dotenv()
    creds_file = 'credentials.json'
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_file(creds_file, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    doc = gc.open_by_key(sheet_id)
    
    for name in ['CEEP_DOPS', 'CEEP_MiniCEX']:
        try:
            ws = doc.worksheet(name)
            headers = ws.row_values(1)
            print(f"Sheet: {name}")
            for i, h in enumerate(headers):
                print(f"  [{i}] {h}")
        except Exception as e:
            print(f"Error reading {name}: {e}")

if __name__ == "__main__":
    get_headers()
