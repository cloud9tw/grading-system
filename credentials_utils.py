import os
import json
import logging
import gspread
from google.cloud import bigquery
from google.oauth2 import service_account

def get_bq_client():
    """
    初始化 BigQuery 客戶端。
    優先順序：1. 本地 credentials.json 2. 環境變數 JSON 字串 3. GCP 環境預設憑證
    """
    try:
        # [策略 1] 優先嘗試從本地檔案讀取 (適用於 Windows/本地開發環境)
        if os.path.exists('credentials.json'):
            credentials = service_account.Credentials.from_service_account_file('credentials.json')
            bq_client = bigquery.Client(credentials=credentials, project=credentials.project_id)
            return bq_client, credentials.project_id
        
        # [策略 2] 針對 Cloud Run 等無檔案環境，從環境變數讀取 JSON 字串內容
        json_str = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON") or os.getenv("GOOGLE_CREDENTIALS_JSON")
        if json_str and json_str.strip().startswith('{'):
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w+', delete=False, suffix='.json') as f:
                f.write(json_str)
                temp_path = f.name
            try:
                credentials = service_account.Credentials.from_service_account_file(temp_path)
                bq_client = bigquery.Client(credentials=credentials, project=credentials.project_id)
                return bq_client, credentials.project_id
            finally:
                if os.path.exists(temp_path): os.remove(temp_path)

        # [策略 3] 最後嘗試使用 GCP Application Default Credentials
        import google.auth
        credentials, project = google.auth.default()
        project = project or os.getenv("GOOGLE_CLOUD_PROJECT") or "epa-grading-system"
        bq_client = bigquery.Client(credentials=credentials, project=project)
        return bq_client, project
    except Exception as e:
        logging.error(f"❌ BigQuery 初始化失敗: {e}")
        return None, None

def get_gspread_client():
    """
    初始化 Google Sheets 客戶端 (gspread)。
    具備多重後援機制，確保在 Cloud Run 上即使檔案遺失也能正常運作。
    """
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    # [策略 1] 優先從環境變數讀取 JSON 字串內容
    json_str = os.getenv("GOOGLE_CREDENTIALS_JSON") or os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if json_str and json_str.strip().startswith('{'):
        try:
            creds_dict = json.loads(json_str)
            return gspread.service_account_from_dict(creds_dict, scopes=scope)
        except Exception as e:
            logging.error(f"❌ 從環境變數 JSON 解析憑證失敗: {e}")
            
    # [策略 2] 嘗試讀取實體檔案
    possible_paths = [
        os.getenv("GOOGLE_APPLICATION_CREDENTIALS"),
        "/etc/secrets/credentials.json",
        "credentials.json",
        os.path.join(os.path.dirname(__file__), 'credentials.json')
    ]
    
    for path in possible_paths:
        if path and os.path.exists(path):
            try:
                return gspread.service_account(filename=path, scopes=scope)
            except Exception as e:
                logging.error(f"❌ 讀取憑證檔案 {path} 失敗: {e}")
                
    # [策略 3] 使用預設憑證
    try:
        import google.auth
        credentials, project = google.auth.default(scopes=scope)
        return gspread.Client(auth=credentials)
    except Exception as e:
        logging.error(f"❌ GCP 預設憑證用於 gspread 失敗: {e}")
                
    raise Exception("無法初始化 Google Sheets 憑證。請確認環境變數或 credentials.json 設定正確。")
