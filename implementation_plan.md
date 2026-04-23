# 🏥 EPA 實習管理系統 - 實作計畫 (Implementation Plan)

本文件詳述系統的階段性開發藍圖，確保各項功能依序實作並與 Google Sheets/BigQuery 架構整合。

## 🎯 專案目標
打造一個自動化、數據驅動的放射實習管理平台，涵蓋從基礎簽到、教學評量到 AI 輔助學習建議的全流程。

---

## 🚀 實作階段規劃

### 🔴 Phase 1: 核心引擎與管理員排程 (已完成 ✅)
1. 實作 `get_current_intern_week(date)` 計算 28 週時間軸。
2. 建立 Google Sheets `學生進度排程` 並透過 API 進行讀寫。
3. 實作管理員端 `schedule_manager.html` 排程維護介面。
4. 實作學生儀表板的「本週進度卡片」。

### 🟡 Phase 2: 進度分析與數據同步 (已完成 ✅)
1. **紀錄比對與儀表板**
   - 系統依據排程自動核對 OPA 評估記錄。
   - 實作「進度達標燈號」（🟢/🔴/⚪）與成長曲線圖表。
2. **BigQuery 同步機制**
   - 實作 Sheets -> BigQuery 的定期數據鏡像，確保大規模數據查詢效能。

### 🔵 Phase 3: 管理員統整報表與匯出 (已完成 ✅)
1. **成績統計匯出**
   - 實作 `aggregate_student_report_data` 整合所有學生數據。
   - 提供 `/api/admin/export_excel` 一鍵下載 Excel 成績總表。
2. **限流保護 (Rate Limiting)**
   - 實作 API 緩衝機制，解決 Google Sheets 429 請求超額問題。

### 🟣 Phase 4: 數據解析精準化與 PDF 匯出 (部分完成 🚧)
1. **KAS 資料解析 (已完成 ✅)**
   - 精確區分 CEEP 表單中的 **數值分數 (KAS)** 與 **質性回饋**，解決欄位錯位問題。
2. **多單位自動同步 (已完成 ✅)**
   - 優化爬蟲以支援教學記錄表單的跨單位重複查詢。
3. **學生履歷 PDF 匯出 (進行中 ⏳)**
   - 整合 `html2pdf.js` 或 `jspdf`，將儀表板圖表與評核統計產出為正式證明文件。

### 🟢 Phase 5: AI 輔助 ILP (Individualized Learning Plan) 雙軌生成
1. **學員數據統整 (Data Aggregation)**
   - 定期彙整每位學員的量性指標與質性回饋。
2. **雙 AI 模型串接比較 (A/B Testing)**
   - 介接 **兩種不同的 AI 模型** (如 GPT-4, Gemini Pro)。
   - 同時產出 ILP 報告，於後台並排展示供教師比對、修改與採納。
3. **週期性產出 (Weekly Routine)**
   - 設定每星期自動觸發生成，建立持續性的學員成長追蹤機制。

---

> 備註：此計畫將隨開發進度動態調整，最新版本儲存於 Git 倉庫。
