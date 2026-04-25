import os
import vertexai
from vertexai.generative_models import GenerativeModel
import logging
from google.cloud import bigquery
from credentials_utils import get_bq_client, get_gspread_client
from privacy_utils import get_code, decode_name
from dotenv import load_dotenv

load_dotenv()

# Initialize Vertex AI
bq_client_obj, project_id = get_bq_client()
if project_id:
    vertexai.init(project=project_id, location="asia-east1")

def fetch_student_performance_data(student_name):
    """
    從 BigQuery 與 Sheets 整合學員表現資料。
    """
    manager_bq, project = get_bq_client()
    student_code = get_code(student_name, 'student')
    
    # 1. 從 BigQuery 抓取 EPA 評分 (使用代碼)
    query = f"""
        SELECT station, opa1_sum, opa2_sum, opa3_sum, comment, aspect1, aspect2, timestamp
        FROM `{project}.grading_data.grading_logs`
        WHERE student_id = '{student_code}'
        ORDER BY timestamp DESC
    """
    epa_records = []
    try:
        query_job = manager_bq.query(query)
        for row in query_job:
            epa_records.append(dict(row))
    except Exception as e:
        logging.error(f"BQ Fetch Error: {e}")

    # 2. 從 CEEP Sheets 抓取質性評語 (使用真實姓名，然後去識別)
    ceep_comments = []
    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        
        for sname in ["CEEP_DOPS", "CEEP_MiniCEX"]:
            try:
                ws = doc.worksheet(sname)
                all_vals = ws.get_all_values()
                if len(all_vals) > 1:
                    headers = all_vals[0]
                    for row in all_vals[1:]:
                        if len(row) > 0 and row[0].strip() == student_name:
                            # 找出評語欄位 (假設在後方，且包含關鍵字)
                            for i, cell in enumerate(row):
                                if len(cell) > 15: # 較長的文字通常是評語
                                    ceep_comments.append({
                                        "source": sname,
                                        "content": cell.strip()
                                    })
            except: continue
    except Exception as e:
        logging.error(f"CEEP Sheets Fetch Error: {e}")

    return {
        "student_code": student_code,
        "epa_records": epa_records,
        "ceep_comments": ceep_comments
    }

def generate_ilp_vertex(student_name):
    """
    調用 Vertex AI (Gemini) 生成 ILP。
    """
    data = fetch_student_performance_data(student_name)
    
    # 建構 Prompt (去識別化)
    prompt = f"你是一位醫學教育專家。請根據以下匿名學員（代號：{data['student_code']}）的評分數據與教師回饋，生成一份專業的個人化學習計畫 (ILP)。\n\n"
    
    prompt += "### 1. EPA 評分紀錄 (分項分數與評語):\n"
    for r in data['epa_records']:
        prompt += f"- 站別: {r['station']} | 分數: {r['opa1_sum']}/{r['opa2_sum']}/{r['opa3_sum']} | 評語: {r['comment']} | 優缺點: {r['aspect1']}, {r['aspect2']}\n"
    
    prompt += "\n### 2. 其他臨床回饋 (質性描述):\n"
    for c in data['ceep_comments']:
        prompt += f"- [{c['source']}]: {c['content']}\n"

    prompt += "\n---要求---\n"
    prompt += "1. 使用繁體中文撰寫。\n"
    prompt += "2. 包含以下結構：## 🎯 優勢摘要、## ⚠️ 待加強領域、## 📅 下一階段學習目標、## 📝 給學員的具體行動建議。\n"
    prompt += "3. 內容需具備不可辨識性，嚴禁出現任何真實姓名。\n"
    prompt += "4. 語氣需專業且具鼓勵性。\n"

    try:
        model = GenerativeModel("gemini-2.0-flash-001")
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        logging.error(f"Vertex AI Error: {e}")
        return f"AI 分析失敗：{str(e)}"

if __name__ == "__main__":
    # 測試
    res = generate_ilp_vertex("測試學生")
    print(res)
