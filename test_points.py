import os
import gspread
from google.oauth2 import service_account
from dotenv import load_dotenv
from gamification import get_leaderboard_data

def test_leaderboard():
    print("Testing Leaderboard Performance & Accuracy...")
    load_dotenv()
    
    creds_file = 'credentials.json'
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_file(creds_file, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    doc = gc.open_by_key(sheet_id)
    
    lb = get_leaderboard_data(gc, doc)
    print("\n--- Current Top 5 ---")
    for i, s in enumerate(lb[:5]):
        print(f"{i+1}. {s['name']}: {s['points']} pts")
    
    if len(lb) > 0 and lb[0]['points'] > 0:
        print("\n✅ Success! Leaderboard data is non-zero.")
    else:
        print("\n⚠️ Warning: Leaderboard points are still 0 or empty.")

if __name__ == "__main__":
    test_leaderboard()
