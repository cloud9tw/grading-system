import os
import gspread
from google.oauth2 import service_account
from dotenv import load_dotenv

def init_settings():
    load_dotenv()
    print("🚀 Initializing System Settings worksheet...")
    try:
        creds_file = 'credentials.json'
        # 1. 直接讀取環境變數中的 JSON 字串 (推薦在 GCP Cloud Run 等無伺服器環境使用)
        # 這裡為了簡單先用實體檔案，邏輯與 app.py 一致
        scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = service_account.Credentials.from_service_account_file(creds_file, scopes=scopes)
        gc = gspread.authorize(creds)
        
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        doc = gc.open_by_key(sheet_id)
        
        # 嘗試新增或讀取
        try:
            ws = doc.add_worksheet(title='系統設定', rows='100', cols='10')
            print("  - Created new worksheet '系統設定'")
        except gspread.exceptions.APIError:
            ws = doc.worksheet('系統設定')
            print("  - Worksheet '系統設定' already exists")
            
        # 更新標頭與初始關鍵字
        ws.update_cell(1, 1, '負面關鍵字')
        
        # 檢查是否已有關鍵字，若無則預填
        if not ws.cell(2, 1).value:
            keywords = [['不耐煩'], ['口氣'], ['態度'], ['不滿'], ['兇'], ['待加強'], ['不滿意']]
            ws.update('A2', keywords)
            print("  ✅ Initial keywords added.")
        else:
            print("  - Keywords already present.")
            
    except Exception as e:
        print(f"  ❌ Initialization failed: {e}")

if __name__ == "__main__":
    init_settings()
