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

### 5. ⏱️ 智能簽到退系統
- **QR Code 辨識**：學員專屬二維碼掃描。
- **API 限流保護**：整合緩衝機制與 BigQuery 快取，徹底解決 Google Sheets API 的 429 限額問題。

---

## 🛠️ 技術架構

- **運算核心**: Flask (Python 3.13) 部署於 **Google Cloud Run**。
- **資料中心**: 
  - **BigQuery**: 主要分析數據庫，負責高頻次數據查詢與聚合分析。
  - **Google Sheets**: 作為設定檔管理與數據鏡像備份。
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
- **Phase 8: AI 個人化學習建議**：利用 Gemini 2.0 針對學員弱點產出成長建議。
