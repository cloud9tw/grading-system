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
- **參數化配置**：計畫名稱、表單標籤與登入帳密已模組化，管理員可於 `ceep_scraper.py` 頂部快速調整。
- **多單位循環抓取**：針對「教學記錄(實習生)」等跨單位表單，支援自動重複查詢並過濾重複項，確保數據完整。
- **API 強化保護**：導入 **指數退避重試機制 (Retry)** 與 **欄位過濾 (Only A:B)**，讀取現有資料時僅抓取必要欄位，大幅降低 Google Sheets API 超時與 429 錯誤機率。

### 4. 🎨 全新登入介面與權限管理
- **頂級視覺設計**：採用橘色漸層底圖與 **毛玻璃 (Glassmorphism)** 效果，提供極致的專業操作體驗。
- **自主權限申請**：提供「想用自己的 Google 帳號登入？」功能，教師可直接填表申請，系統將自動寄信通知管理員。

### 5. 🏆 英雄榜與 Gamification
- **即時排行**：基於 BigQuery 數據計算，呈現學員表現積分與英雄榜成就。

---

## 🛠️ 技術架構

- **運算核心**: Flask (Python 3.13) 部署於 **Google Cloud Run**。
- **資料中心**: 
  - **BigQuery**: 主要分析數據庫，驅動儀表板、成長曲線與英雄榜。
  - **Google Sheets**: 作為設定管理與數據鏡像備份，提供管理員直觀的維護介面。
- **AI 引擎**: 整合 **GCP Vertex AI (Gemini 2.0 Flash)**，直接利用現有服務帳號進行認證。
- **視覺識別**: 整合自訂 ICON 系統 (橘底紫色盾牌圖案)，提升系統專業感。

---

## ⚙️ 快速上手

### 1. 安裝與環境
```powershell
pip install -r requirements.txt
playwright install chromium
```

### 2. 環境變數 (.env)
需設定 GCP 憑證路徑及相關 Email 通知帳密。
```env
GOOGLE_SERVICE_ACCOUNT_JSON=credentials.json
SENDER_EMAIL=...
SENDER_PASSWORD=...
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

---

## 🚀 未來規劃 (Phase 5: AI ILP)

本專案現正進入 AI 驅動的教學優化階段：
- **AI 個人化學習計畫 (ILP)**：彙整學員量性指標與質性評語，透過 **GCP Vertex AI** 產出 ILP。
- **匿名化分析技術**：在資料傳送至 AI 前，系統會自動進行「去識別化」，確保在完全匿名（不可辨識性）的前提下生成專業建議。

詳細開發藍圖請參閱 [implementation_plan.md](implementation_plan.md)。
