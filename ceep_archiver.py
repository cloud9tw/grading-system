import os
import gspread
from dotenv import load_dotenv

load_dotenv()

from credentials_utils import get_gspread_client

def archive_to_sheets(records, sheet_name="CEEP_DOPS"):
    """
    Append list of records to the specified worksheet in Google Sheets.
    """
    if not records:
        print(f"[{sheet_name}] 沒有任何紀錄需要存入。")
        return

    gc = get_gspread_client()
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    doc = gc.open_by_key(sheet_id)

    # 偵測最大的評分項目數量
    max_items = 0
    for rec in records:
        item_count = len(rec.get("scores", {}))
        if item_count > max_items:
            max_items = item_count
    
    # 基本欄位
    headers = ["學員姓名", "送出時間", "個案名稱", "計畫/開始時間"]
    for i in range(1, max_items + 1):
        headers.append(f"項目_{i}")

    def get_friendly_error(e):
        err_msg = str(e)
        if "429" in err_msg or "Quota exceeded" in err_msg:
            return "資料讀取達到上限，請稍候數秒再試 (Google Sheets API 限制)"
        return err_msg

    import time
    # 1. 取得或建立工作表
    worksheet = None
    for attempt in range(3):
        try:
            worksheet = doc.worksheet(sheet_name)
            break
        except gspread.exceptions.WorksheetNotFound:
            worksheet = doc.add_worksheet(title=sheet_name, rows="1000", cols=str(len(headers)))
            worksheet.append_row(headers)
            print(f"✅ 已建立新分頁: {sheet_name}")
            break
        except Exception as e:
            friendly = get_friendly_error(e)
            print(f"⚠️ 嘗試開啟工作表失敗 ({attempt+1}/3): {friendly}")
            time.sleep(2)
    
    if not worksheet:
        raise Exception("資料讀取達到上限，請稍候數秒再試" if worksheet is None else f"無法存取工作表: {sheet_name}")

    # 2. 獲取現有資料以避免重複
    unique_keys = set()
    for attempt in range(3):
        try:
            existing_keys_data = worksheet.get('A:B')
            if len(existing_keys_data) > 1:
                for row in existing_keys_data[1:]:
                    if len(row) >= 2:
                        unique_keys.add((row[0].strip(), row[1].strip()))
            break
        except Exception as e:
            friendly = get_friendly_error(e)
            print(f"⚠️ 讀取現有資料失敗 ({attempt+1}/3): {friendly}")
            if attempt == 2: raise Exception("資料讀取達到上限，請稍候數秒再試")
            time.sleep(3)

    # 3. 準備寫入資料
    new_rows = []
    for rec in records:
        key = (rec["student_name"], rec["submit_time"])
        if key not in unique_keys:
            row = [
                rec["student_name"],
                rec["submit_time"],
                rec["case_name"],
                rec["start_time"]
            ]
            for i in range(1, max_items + 1):
                row.append(rec["scores"].get(f"item_{i}", ""))
            new_rows.append(row)
            unique_keys.add(key)

    # 4. 批次寫入
    if new_rows:
        for attempt in range(3):
            try:
                worksheet.append_rows(new_rows)
                print(f"✅ [{sheet_name}] 成功歸檔 {len(new_rows)} 筆新數據")
                break
            except Exception as e:
                friendly = get_friendly_error(e)
                print(f"⚠️ 批次寫入失敗 ({attempt+1}/3): {friendly}")
                if attempt == 2: raise Exception("資料讀取達到上限，請稍候數秒再試")
                time.sleep(4)
    else:
        print(f"ℹ️ [{sheet_name}] 所有資料皆已存在，無需更新。")

if __name__ == "__main__":
    # Test with dummy data
    test_records = [{
        "student_name": "測試學生",
        "submit_time": "2024-04-21 10:00:00",
        "case_name": "測試個案",
        "start_time": "2024-04-21 09:00:00",
        "scores": {"item_1": "通過", "item_2": "優"}
    }]
    archive_to_sheets(test_records)
