import os
import gspread
import logging
from dotenv import load_dotenv
from credentials_utils import get_gspread_client

load_dotenv()

class PrivacyManager:
    """
    管理姓名與代碼之間的對照 (Anonymization mapping)。
    對照表儲存在 Google Sheets 中的『系統隱私查照表』分頁。
    """
    SHEET_TITLE = "系統隱私查照表"
    HEADERS = ["原始姓名", "類別", "匿名代碼"]

    def __init__(self):
        self.gc = get_gspread_client()
        self.sheet_id = os.getenv("GOOGLE_SHEET_ID")
        self.doc = self.gc.open_by_key(self.sheet_id)
        self._cache = {} # {(name, type): code}
        self._reverse_cache = {} # {code: name}
        self._worksheet = None
        self._load_worksheet()

    def _load_worksheet(self):
        try:
            self._worksheet = self.doc.worksheet(self.SHEET_TITLE)
        except gspread.exceptions.WorksheetNotFound:
            # 自動建立分頁
            self._worksheet = self.doc.add_worksheet(title=self.SHEET_TITLE, rows="1000", cols="3")
            self._worksheet.append_row(self.HEADERS)
            logging.info(f"✅ 已建立隱私對照表分頁: {self.SHEET_TITLE}")
        
        # 載入現有對照資料到快取
        records = self._worksheet.get_all_records()
        for r in records:
            name = r["原始姓名"]
            ctype = r["類別"]
            code = r["匿名代碼"]
            self._cache[(name, ctype)] = code
            self._reverse_cache[code] = name

    def get_code(self, name, ctype="student"):
        """
        將姓名轉換為代碼。若無記錄則自動生成並更新至 Sheets。
        ctype: 'student' 或 'teacher'
        """
        name = str(name).strip()
        if not name: return ""
        
        # 檢查快取
        key = (name, ctype)
        if key in self._cache:
            return self._cache[key]

        # 生成新代碼
        prefix = "S" if ctype == "student" else "T"
        # 計算現有該類別的數量
        existing_codes = [c for (n, t), c in self._cache.items() if t == ctype]
        next_num = len(existing_codes) + 1
        new_code = f"{prefix}{next_num:04d}" # 例如 S0001, T0001

        # 防止極其罕見的重複 (雖然機率極低)
        while new_code in self._reverse_cache:
            next_num += 1
            new_code = f"{prefix}{next_num:04d}"

        # 更新至 Sheets 與快取
        try:
            self._worksheet.append_row([name, ctype, new_code])
            self._cache[key] = new_code
            self._reverse_cache[new_code] = name
            logging.info(f"🆕 已為 {ctype} '{name}' 生成新代碼: {new_code}")
            return new_code
        except Exception as e:
            logging.error(f"❌ 無法更新隱私對照表: {e}")
            return f"ERR_{new_code}" # 備援

    def decode_name(self, code):
        """
        將代碼還原為姓名。
        """
        return self._reverse_cache.get(code, code)

# Singleton 實例
_manager = None

def get_privacy_manager():
    global _manager
    if _manager is None:
        _manager = PrivacyManager()
    return _manager

def get_code(name, ctype="student"):
    return get_privacy_manager().get_code(name, ctype)

def decode_name(code):
    return get_privacy_manager().decode_name(code)

if __name__ == "__main__":
    # 簡單測試
    print(f"Student: {get_code('王小明', 'student')}")
    print(f"Teacher: {get_code('林教授', 'teacher')}")
    print(f"Decode: {decode_name('S0001')}")
