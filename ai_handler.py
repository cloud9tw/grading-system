import os
import logging
from openai import OpenAI
from google.cloud import bigquery
from credentials_utils import get_bq_client, get_gspread_client
from privacy_utils import get_code, decode_name
from dotenv import load_dotenv

load_dotenv()

# ==========================================
# 🧠 AI 固定指令設定 (Fixed Prompt)
# 您可以在這裡修改分析的風格、重點或格式
# ==========================================
SYSTEM_PROMPT = """你是一位專業的臨床醫學教育專家。
你的任務是分析學員的臨床評分數據（EPA/DOPS/Mini-CEX）與教師的質性評語，
並生成一份結構化的「個人化學習計畫 (Individual Learning Plan, ILP)」。

要求：
1. 使用繁體中文。
2. 語氣專業、嚴謹且具有教育指導意義。
3. 嚴禁出現任何學員或教師的真實姓名，僅使用代號。
4. 結構必須包含：
   - ## 🎯 核心能力摘要
   - ## 🔍 表現優勢分析
   - ## ⚠️ 臨床操作盲點與改進空間
   - ## 🚀 具體行動建議 (依站別列出)
   - ## 📈 學習進度預期
"""
# ==========================================

# 初始化 OpenAI Client
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def fetch_student_performance_data(student_name):
    """
    從 BigQuery 與 Sheets 整合學員表現資料。
    """
    manager_bq, project = get_bq_client()
    student_code = get_code(student_name, 'student')
    
    # 1. 從 BigQuery 抓取 EPA 評分
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

    # 2. 從 CEEP Sheets 抓取質性評語
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
                    for row in all_vals[1:]:
                        if len(row) > 2 and row[2].strip() == student_name: # 假設姓名在第三欄
                            comment_content = ""
                            # 尋找內容較長的欄位作為評語
                            for cell in row:
                                if len(str(cell)) > 15:
                                    comment_content = str(cell).strip()
                                    break
                            if comment_content:
                                ceep_comments.append({"source": sname, "content": comment_content})
            except: continue
    except Exception as e:
        logging.error(f"CEEP Sheets Fetch Error: {e}")

    return {
        "student_code": student_code,
        "epa_records": epa_records,
        "ceep_comments": ceep_comments
    }

def generate_ilp_chatgpt(student_name):
    """
    調用 OpenAI ChatGPT 生成 ILP。
    """
    data = fetch_student_performance_data(student_name)
    
    # 建構使用者資料 Prompt
    user_content = f"學員代號：{data['student_code']}\n\n"
    user_content += "### 數據紀錄：\n"
    for r in data['epa_records']:
        user_content += f"- [{r['station']}] 分數: {r['opa1_sum']}/{r['opa2_sum']}/{r['opa3_sum']} | 評語: {r['comment']} | 亮點: {r['aspect1']}, 改善點: {r['aspect2']}\n"
    
    user_content += "\n### 額外評語：\n"
    for c in data['ceep_comments']:
        user_content += f"- {c['source']}: {c['content']}\n"

    try:
        response = client.chat.completions.create(
            model="gpt-4o", # 建議使用 gpt-4o 或 gpt-4o-mini
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_content}
            ],
            temperature=0.7
        )
        return response.choices[0].message.content
    except Exception as e:
        logging.error(f"OpenAI Error: {e}")
        return f"ChatGPT 分析失敗：{str(e)}"

if __name__ == "__main__":
    # 測試 (需先在環境變數設定 OPENAI_API_KEY)
    res = generate_ilp_chatgpt("測試學生")
    print(res)
