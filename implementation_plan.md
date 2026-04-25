# 🏥 EPA 系統開發進度與實作計畫 (Updated 2026-04-24)

## 🟢 Phase 1-3: 基礎建設與數據整合 (已完成 ✅)

- **站別評分表單**：實作 OPA1, 2, 3 數據收集與 Google Sheets 儲存。
- **QR 簽到系統**：實作學員二維碼簽到退與自動小時數計算。
- **CEEP 同步**：實作爬蟲自動抓取 DOPS / Mini-CEX 數據。

## 🔵 Phase 4-6: 數據引擎、AI 分析與隱私保護 (已完成 ✅)

- **BigQuery 數據引擎 (全面化)**：
  - [DONE] 遷移學員統計、CEEP DOPS/MiniCEX 數據至 BQ。
  - [DONE] 實作 BQ 優先寫入策略，解決 Sheets API 429 寫入失敗問題。
- **去識別化與隱私管理**：
  - [DONE] 實作 `PrivacyManager` 將 BQ 資料代碼化（SXXXX/TXXXX）。
  - [DONE] 建立 Google Sheets 隱私對照表自動同步機制。
- **AI 個人化學習計畫 (ILP)**：
  - [DONE] 整合 Vertex AI (Gemini 2.0 Flash) 生成學習建議報告。
  - [DONE] 實作管理員 AI 分析中心與 Markdown 報告預覽介面。
- **Pro 版專業成長報表 & 自動化 PDF**：
  - [DONE] 實作多選站別動態過濾與身分感應參考線。
  - [DONE] 實作分站批次 PDF 產出功能。

## 🟡 Phase 7: 會議前置評量與分發系統 (進行中 ⏳)

- **[NEW] 會議準備中心**：一次選取多位學員，自動發送報表連結給教師。
- **[NEW] 預先評核系統（依職級動態限制）**：
  - 等級清單：1, 2a, 2b, 3a, 3b, 3c, 4, 5（固定清單）。
  - **實習生**：最高選至 **3b**，3c 以上選項自動停用（灰化）。
  - **PGY**：可選至完整 **5** 級。
  - 前端依 `student.type` 動態渲染可選等級，防止評分超出預期目標。
  - 獨立儲存於 `pre_meeting_evals` 資料表（欄位：`student_id`, `meeting_date`, `evaluator_email`, `level`, `note`）。
- **[NEW] 預評分數整合顯示**：將彙整後的預評等級與評語，同步顯示於專業版成長報告中，作為會議討論的「預期達成基準」。
- **[NEW] 郵件通知自動化**：整合 SMTP/API 發送報表閱讀邀請。

## 🟣 Phase 8: AI 個人化教學優化 (已完成 ✅)
- **ChatGPT 引擎導入**：
  - [DONE] 實作 `ai_handler.py` 整合 OpenAI GPT-4o。
  - [DONE] 支援「系統固定 Prompt」設定，確保分析風格一致。
  - [DONE] 自動彙整 BQ 量化指標與 CEEP 質性評語。
- **資料隱私安全層**：
  - [DONE] 實作 `privacy_utils.py` 與 `privacy_mapping`。
  - [DONE] 實作 AI 分析前自動去識別化流程，嚴禁真實姓名進入 AI 模型。
  - [DONE] 實作 ILP 報告之代碼自動還原顯示。

## 🟢 Phase 9: 系統維護與效能優化 (持續中 ⏳)
- **全域快取保護**：已擴充至遊戲化數據與設定檔。
- **BigQuery 授權統一**：修復本地與雲端憑證衝突問題。
- **錯誤監控**：加入 429 配額預警與 API 健康檢查。

---

> 備註：此計畫根據臨床教學需求動態調整。最新代碼已同步至 GitHub `main` 分支。
