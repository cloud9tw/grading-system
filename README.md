# 🏥 EPA 臨床實習評分與出勤管理系統

![Python](https://img.shields.io/badge/Python-3.13-blue)
![Flask](https://img.shields.io/badge/Framework-Flask-green)
![GCP](https://img.shields.io/badge/Deployment-CloudRun-blue)
![BigQuery](https://img.shields.io/badge/Database-BigQuery-coral)
![GoogleSheet](https://img.shields.io/badge/Storage-GoogleSheets-yellow)

本系統是一套專為醫事放射職類開發的臨床教學評選工具，採 **BigQuery + Google Sheets 雙軌架構**。提供即時評分、QR Code 簽到退、自動化成績彙整以及 CEEP 外部數據同步等全方位功能，並已優化為生產環境等級的 Cloud Run 部署。

---

## 🌟 核心功能

### 1. 📝 臨床教學評分 (EPA Scoring)
- **站別化評核**：支援 CT、MR、ROUTINE 等多種檢查站別。
- **OPA 階段性評價**：細分為 OPA1 (前準備)、OPA2 (執行)、OPA3 (後處置) 三大部分。
- **質性與量化並行**：支援 8 級信賴度量表 (Entrustment Levels) 與 15 字以上質性評論。

### 2. ⏱️ 智能簽到退系統 (Attendance Control)
- **QR Code 辨識**：學員專屬二維碼掃描，確保出勤真實性。
- **自動預警系統**：針對遲到、早退自動寄送 Email 給管理員。

### 3. 🔄 CEEP 資料同步與精準解析 (Advanced Data Sync)
- **多單位循環抓取**：針對「教學記錄(實習生)」等跨單位表單，支援自動重複查詢 (預設 3 次) 確保數據完整。
- **KAS 與回饋分離**：修正 CEEP 表單格式差異問題，自動區分 **數值分數 (KAS)** 與 **質性回饋 (Qualitative Feedback)**，避免報表欄位錯位。
- **API 配額保護**：在 Sheets 讀寫邏輯中導入 `time.sleep(1)` 緩衝，徹底解決 Google Sheets API 429 (Quota Exceeded) 錯誤。

### 4. 🏆 英雄榜與 Gamification
- **即時排行**：基於 BigQuery 數據計算，呈現學員表現積分與英雄榜成就。

---

## 🛠️ 技術架構

- **運算核心**: Flask (Python 3.13) 部署於 **Google Cloud Run**。
- **資料中心**: 
  - **BigQuery**: 主要分析數據庫，驅動儀表板、成長曲線與英雄榜。
  - **Google Sheets**: 作為設定管理與數據鏡像備份，提供管理員直觀的維護介面。
- **視覺識別**: 整合自訂 ICON 系統 (橘底紫色盾牌圖案)，提升系統專業感。

---

## ⚙️ 快速上手

### 1. 安裝與環境
```powershell
pip install -r requirements.txt
playwright install chromium
```

### 2. 部署至 Cloud Run
```powershell
gcloud run deploy epa-grading-system --source . --region asia-east1
```

---

## 📊 成績計算與報表邏輯 (Updated)

報表彙整邏輯位於 `app.py` 的 `aggregate_student_report_data`：

### 1. OPA 成績 (分站別)
- 透過 BigQuery 聚合計算各檢查室的 `OPA1+OPA2+OPA3` 平均分。

### 2. DOPS / Mini-CEX (分站別精確解析)
- **DOPS**：抓取 `CEEP_DOPS` 索引 `[26]` 作為分數，`[23]` 作為老師回饋。
- **Mini-CEX**：抓取 `CEEP_MiniCEX` 索引 `[21]` 作為分數，`[20]` 作為老師回饋。
- **輸出格式**：數值與回饋在 Excel 中將各歸其位，不再混雜於同一欄。

### 3. 進度達標率
- 公式：`(已達標週數 / 目前實習進度週數) * 100%`。

---

## 🚀 未來規劃 (Phase 5)

本專案正朝向 AI 輔助教學目標邁進：
- **AI 雙軌 ILP 生成**：彙整學員量性指標 (Scores) 與質性指標 (Feedback)，每週透過兩種 AI 模型 (如 GPT-4 / Gemini) 雙軌產出個人化學習計畫 (Individualized Learning Plan, ILP) 供教師參考。

詳細開發藍圖請參閱 [implementation_plan.md](implementation_plan.md)。
