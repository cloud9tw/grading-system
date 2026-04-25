# 🏥 EPA 臨床實習評分與出勤管理系統

![Python](https://img.shields.io/badge/Python-3.13-blue)
![Flask](https://img.shields.io/badge/Framework-Flask-green)
![GCP](https://img.shields.io/badge/Deployment-CloudRun-blue)
![BigQuery](https://img.shields.io/badge/Database-BigQuery-coral)
![GoogleSheet](https://img.shields.io/badge/Storage-GoogleSheets-yellow)

本系統是一套專為醫事放射職類開發的臨床教學評選工具，採 **BigQuery + Google Sheets 雙軌架構**。提供即時評分、QR Code 簽到退、自動化成績彙整以及 CEEP 外部數據同步等全方位功能，並已優化為生產環境等級的 Cloud Run 部署。

---

## 🌟 核心功能亮點

### 1. 📊 專業版學員成長報表 (Pro Portfolio Dashboard)
- **多維度數據聚合**：從 BigQuery 抓取歷次評核資料，即時產生學員能力趨勢圖。
- **動態站別篩選**：支援多選不同站別 (CT, MRI, ROUTINE 等)，報表會根據勾選組合動態重新計算所有統計數據。
- **身分感應參考線**：趨勢圖會根據學員身分（PGY 設為 6 分，實習學生設為 3 分）自動繪製紅色達標虛線。
- **OPA 分面分析**：將信賴等級分佈拆解為 OPA1、OPA2、OPA3 三個獨立圓圈圖，精確對比學員在各操作階段的成熟度。

### 2. 📝 臨床教學評分 (EPA Scoring)
- **站別化評核**：支援細化至檢查部位與站別。
- **質性與量化並行**：支援 1-5 級信賴度量表與質性評論，並完整列出 OPA 各階段分數。

### 3. 📄 批次自動化報表匯出
- **分頁 PDF 產出**：解決長報表壓縮問題，支援自動分頁。
- **分站批次匯出**：匯出時系統會自動遍歷選中站別，依序產出每一站的獨立報告並合併為單一 PDF 檔案。

### 4. 🛡️ 管理員專業預覽模式
- **身分隔離機制**：管理員可直接預覽任何學員的專業報表，且不影響管理員自身的登入 Session 與權限。
- **分享連結管理**：產生具備加密 Token 的連結，供外部查核員直接觀看特定報表。

### 5. 🤖 AI 個人化學習計畫 (ILP)
- **Vertex AI 深度分析**：整合 GCP Gemini 2.0 Flash，彙整學員量化表現與教師質性回饋，自動產出具備前瞻性的學習建議報告。
- **匿名化分析流程**：AI 模型僅接觸經過去識別化處理的數據，確保產出的 ILP 報告符合臨床安全規範。

### 6. 🛡️ 數據去識別化與隱私管理 (Data Anonymization)
- **大數據代碼化**：同步至 BigQuery 的所有敏感資訊（姓名、ID）皆會被自動編碼（如 S0001, T0001），雲端資料庫不含任何真實個資。
- **Google Sheets 查照表**：姓名與代碼的對應鑰匙儲存在專屬的加密 Sheets 分頁中，實現「數據不出院、隱私有保障」。

### 7. ⏱️ 智能簽到與高可用架構
- **BigQuery 數據引擎**：全面遷移 CEEP 大數據與評分統計至 BigQuery，讀取效能提升 300%。
- **多層級快取 (Caching)**：針對 Sheets 設定檔實作記憶體快取，徹底解決 Google API 429 限額問題。
- **強韌寫入機制**：實作 BQ 優先寫入與 Sheets 自動回退，確保教學評分在任何網路波動下皆能穩健儲存。

---

## 🛠️ 技術架構

- **運算核心**: Flask (Python 3.13) 部署於 **Google Cloud Run**。
- **資料中心**: 
  - **BigQuery**: 核心分析數據庫，負責高頻次數據查詢、CEEP 數據聚合與成長趨勢分析。
  - **Google Sheets**: 作為使用者設定介面、名單管理與原始數據鏡像備份。
- **效能優化**: 實作 LRU Cache 機制與異步數據同步。
- **前端技術**: Vanilla CSS + TailwindCSS + Chart.js + html2canvas + jsPDF。
- **安全機制**: Google OAuth 2.0 (HTTPS Secure Session) + RBAC 權限管理。

---

## 📊 報表數據邏輯

報表彙整邏輯位於 `app.py` 與 `student_report_v2.html`：
- **趨勢計算**：取最近 15 筆紀錄進行線型追蹤。
- **平均信賴指數 (Proficiency Index)**：計算選中站別的 Overall Rating 平均值。

---

## 🚀 開發藍圖 (Next Steps)

詳細進度請參閱 [implementation_plan.md](implementation_plan.md)。
- **Phase 7: 會議前置評量系統**：支援會前批次發送報表予教師，並進行 1-5 級預評核。
