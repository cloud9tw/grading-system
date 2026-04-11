import os

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

old_attn = '''    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        sheet = doc.worksheet('上下班打卡記錄')
        all_vals = sheet.get_all_values()
        
        target_name = student_info['name']
        history = []
        
        for r in all_vals[1:]:
            # Header: ['學生', '教師', '共同教師', '檢查室', '簽到時間', '簽退時間']
            if len(r) >= 5 and str(r[0]).strip() == target_name:
                check_in_time = str(r[4]).strip()
                check_out_time = str(r[5]).strip() if len(r) > 5 else ''
                history.append({
                    'room': str(r[3]).strip(),
                    'check_in': check_in_time,
                    'check_out': check_out_time,
                    'is_complete': bool(check_in_time and check_out_time)
                })
                
        # 依簽到時間反查排序 (最新的在前)
        history.sort(key=lambda x: x['check_in'], reverse=True)
        
        return jsonify({'success': True, 'data': history})'''

new_attn = '''    try:
        target_name = student_info['name']
        history = []
        bq_client, project_id = get_bq_client()
        if bq_client:
            q = f"SELECT sub_room, check_in_time, check_out_time FROM \{project_id}.grading_data.attendance_daily_summary\ WHERE student_name = @sname ORDER BY check_in_time DESC"
            job_config = bq_client.query_class.QueryJobConfig(
                query_parameters=[bq_client.query_class.ScalarQueryParameter("sname", "STRING", target_name)]
            )
            res = bq_client.query(q, job_config=job_config).result()
            
            for r in res:
                history.append({
                    'room': r.sub_room,
                    'check_in': r.check_in_time.strftime('%Y/%m/%d %H:%M:%S') if r.check_in_time else '',
                    'check_out': r.check_out_time.strftime('%Y/%m/%d %H:%M:%S') if r.check_out_time else '',
                    'is_complete': bool(r.check_in_time and r.check_out_time)
                })
        
        return jsonify({'success': True, 'data': history})'''

content = content.replace(old_attn, new_attn)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print('Done app.py attn')
