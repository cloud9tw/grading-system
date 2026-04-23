# 🏥 EPA 臨床實習評分與出勤管理系統

![Python](https://img.shields.io/badge/Python-3.13-blue)
![Flask](https://img.shields.io/badge/Framework-Flask-green)
![GoogleSheet](https://img.shields.io/badge/Database-GoogleSheets-yellow)
![Playwright](https://img.shields.io/badge/Automation-Playwright-orange)

本系統是一套專為醫事放射職類開發的臨床教學評選工具，整合了 Google Sheets 作為雲端資料庫，提供即時評分、QR Code 簽到退以及 CEEP 外部數據自動同步等全方位功能。

---

## 🌟 核心功能

### 1. 📝 臨床教學評分 (EPA Scoring)
- **站別化評核**：支援 CT、MR、ROUTINE 等多種檢查站別。
- **OPA 階段性評價**：細分為 OPA1 (前準備)、OPA2 (執行)、OPA3 (後處置) 三大部分。
- **強制回饋機制**：實作「綜合意見 5 選 2」強制勾選，確保教學回饋的深度。
- **質性與量化並行**：支援 8 級信賴度量表 (Entrustment Levels) 與 15 字以上質性評論。

### 2. ⏱️ 智能簽到退系統 (Attendance Control)
- **QR Code 辨識**：學員專屬二維碼掃描，確保出勤真實性。
- **自動預警系統**：
  - **遲到警報**：簽到晚於 08:40 觸發。
  - **早退警報**：簽退早於 17:00 觸發。
  - **Email 通知**：異常發生時，系統會自動發信給設定的管理員 (Notify Emails)。

### 3. 🏆 學員英雄榜與積分 (Gamification)
- **成就獎章**：根據實習表現與簽到率自動加總積分。
- **動態排行**：即時顯示前三名「王者」與「獎章」獲得數，提升實習競爭力與動力。

### 4. 🔄 CEEP 資料同步 (Multi-Form Data Sync)
- **多表單自動化 Scraping**：整合 Playwright 技術，一鍵登入 `ceep2.tmu.edu.tw` 高效同步數據。
- **支援表單類型**：
  - **DOPS 評定**：醫學影像技術學-操作技能直接觀察評量表。
  - **Mini-CEX 評量**：醫學影像技術學-迷你臨床演練評量表。
- **自動化歸檔與美化**：
  - **自動建立分頁**：若雲端無對應分頁，系統會自動建立 `CEEP_DOPS` 與 `CEEP_MiniCEX`。
  - **動態欄位適應**：同步過程會自動偵測評分項次數量，動態生成試算表欄位，確保數據完整不遺漏。

---

## 🛠️ 技術架構

- **後端**: Flask (Python 3.13)
- **前端**: Vanilla CSS (Premium Dark Mode Design), JavaScript
- **身份驗證**: Google OAuth 2.0 (Authlib)
- **資料庫**: 
  - **BigQuery (核心)**: 主要數據儲存體，驅動學員儀表板、成長曲線與英雄榜。
  - **Google Sheets (鏡像)**: 臨床評分紀錄備份與設定檔管理，方便管理員隨時閱覽與手動修正。
- **自動化**: Playwright (CEEP 數據爬蟲)、BigQuery 自動同步排程

---

## ⚙️ 快速上手

### 1. 下載與安裝
```powershell
git clone https://github.com/cloud9tw/grading-system.git
cd grading-system
pip install -r requirements.txt
playwright install chromium
```

### 2. 環境變數設定 (.env)
請在根目錄建立 `.env` 檔案並填入以下內容：
```env
FLASK_SECRET_KEY=您的加密金鑰
GOOGLE_CLIENT_ID=OAuth客戶端ID
GOOGLE_CLIENT_SECRET=OAuth客戶端密鑰
GOOGLE_SHEET_ID=1RlYuWGG8lMjNiL447swS2ZvsWO8oqvc221THi4mcb0I
SENDER_EMAIL=發信用的Gmail
SENDER_PASSWORD=Gmail應用程式密碼
NOTIFY_EMAILS=接收通知的人員Email
```

### 3. 啟動伺服器
```powershell
python app.py
```

---

## 📂 Google Sheets 資料結構說明

系統會讀寫以下工作表，請勿隨意修改表頭：
- **`教師名單`**: 定義 `教師_Email` 與 `管理員權限` (填入 `admin` 即可使用同步功能)。
- **`學員名單`**: 存放學員 `Email`、`學生ID` 與基本資料。
- **`Evaluations`**: 存放所有系統內的評分紀錄。
- **`上下班打卡記錄`**: 存放所有簽到退數據。
- **`CEEP_DOPS`**: (自動生成) 存放從 CEEP 同步過來的 DOPS 數據。
- **`CEEP_MiniCEX`**: (自動生成) 存放從 CEEP 同步過來的 Mini-CEX 數據。
- **`系統設定`**: 全球系統參數與預警設定。
  - **A 欄 (負面關鍵字)**: 若教學回饋命中此處詞彙，會自動發信給管理員。
  - **B 欄 (排除日期)**: 指定不寄送缺席警報的日期 (格式範例：`2024-05-01`)。

---

## 🛡️ 安全與權限
- **教師權限**: 可填寫評分表與查看英雄榜。
- **學員權限**: 僅能執行掃描簽到退與填寫教學回饋 (`/feedback`)。
- **系統管理員**: 教師名單中權限設為 `admin` 者，可進入「管理中心」管理排除日期並手動觸發 CEEP 同步。

## 🎓 階段性評量要求說明

系統會根據《學員名單》中的「學員類別」自動套用不同的評量門檻。這些門檻定義於 Google Sheets 的 **`各類別EPA需求`** 工作表中。

### 📌 評量基準範例
| 學員階段 | EPA/OPA 需求數量 | DOPS / Mini-CEX | 教學回饋表需求 |
| :--- | :--- | :--- | :--- |
| **一般實習生** | 基礎項目 10-15 筆 | 至少各 1 筆 | 每個檢查室 1 筆 |
| **一年制學員 (R1)** | 進階站別 20+ 筆 | 至少各 3 筆 | 每個檢查室 2 筆 |
| **二年制學員 (R2)** | 獨立執行 30+ 筆 | 至少各 5 筆 | 每個檢查室 3 筆 |

---

## 🛠️ 管理員維護指南
1. **修改評量目標**：若需調整各階段所需數量，請至 `各類別EPA需求` 試算表修改對應欄位數值，排行榜積分與進度條會即時更新。
2. **手動同步 CEEP**：在管理入口點擊「CEEP 數據自動同步」，系統會開啟背景任務抓取最新分數並自動歸檔至對應分頁。
3. **匯出成績總表**：在管理中心點擊「匯出實習成績總表」，系統會統整所有數據並產出 Excel 報表。

---

## 📊 成績計算與報表邏輯

報表匯出邏輯位於 `app.py` 的 `aggregate_student_report_data` 函式中，方便管理員隨時微調公式：

### 1. 實習總時數
- **來源**：`上下班打卡記錄` 工作表。
- **計算**：`簽退時間 - 簽到時間` 之總和（單位：小時）。

### 2. OPA 成績 (分站別)
- **來源**：BigQuery `grading_logs` 原始數據。
- **計算**：依據「站別」分類，計算該生在各檢查室的 `OPA1+OPA2+OPA3` 總評平均分。

### 3. DOPS / Mini-CEX (分站別)
- **來源**：`CEEP_DOPS` 與 `CEEP_MiniCEX` 工作表。
- **計算**：依據表單內的「檢查項目/站別」進行分類，提取表單最後一欄的總結得分進行平均計算。

### 4. 進度達標率
- **公式**：`(已達標週數 / 目前實習進度週數) * 100%`。
- **達標判定**：該週所排定的所有站別，於當週日期區間內皆有至少一筆 OPA 評核紀錄。

---
