import os
import json
import traceback
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials

load_dotenv()

def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_file = os.getenv("GOOGLE_APPLICATION_CREDENTIALS") or os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    print(f"Using creds_file: {creds_file}")
    if creds_file and os.path.exists(creds_file):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
            client = gspread.authorize(creds)
            return client
        except Exception as e:
            print(f"Error loading credentials: {e}")
            return None
    else:
        print(f"Credentials file not found or path not set: {creds_file}")
    return None

def test():
    try:
        gc = get_gspread_client()
        if not gc:
            print("Google Sheets backend not configured.")
            return
            
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        print(f"Opening sheet ID: {sheet_id}")
        doc = gc.open_by_key(sheet_id)
        
        print("Worksheets available in this document:")
        for ws in doc.worksheets():
            print(f"- {ws.title}")
            
        print("Trying to open '學員名單'...")
        sheet_students = doc.worksheet('學員名單')
        print("Success for 學員名單")
        
        print("Trying to open '站別OPA細項'...")
        sheet_stations = doc.worksheet('站別OPA細項')
        print("Success for 站別OPA細項")

        print("Trying to open '信賴等級描述及轉換'...")
        sheet_trust = doc.worksheet('信賴等級描述及轉換')
        print("Success for 信賴等級描述及轉換")
        
        # Test headers logic
        st_rec = sheet_students.get_all_records()
        print("Students records:", len(st_rec))
        
        st_opa = sheet_stations.get_all_records()
        print("Stations records:", len(st_opa))
        
        st_trust = sheet_trust.get_all_records()
        print("Trust records:", len(st_trust))
        
        print("All parsing looks good.")
        
    except Exception as e:
        print(f"Exception encountered: {type(e).__name__} -> {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    test()
