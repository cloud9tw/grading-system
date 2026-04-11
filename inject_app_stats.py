import os

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

old_stats = '''    try:
        gc = get_gspread_client()
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        doc = gc.open_by_key(sheet_id)
        sheet = doc.worksheet('評分記錄')
        all_vals = sheet.get_all_values()
        
        records_by_station = {}
        target_id = student_info['id']
        
        for r in all_vals[1:]:
            if len(r) > 8 and str(r[0]).strip() == target_id:
                station = str(r[2]).strip()
                time_str = str(r[4]).strip()
                opa1 = str(r[6]).strip()
                opa2 = str(r[7]).strip()
                opa3 = str(r[8]).strip()
                
                def to_int(v):
                    try: return int(v)
                    except: return None
                        
                pt = {
                    'time': time_str,
                    'opa1': to_int(opa1),
                    'opa2': to_int(opa2),
                    'opa3': to_int(opa3)
                }
                
                if station not in records_by_station:
                    records_by_station[station] = []
                records_by_station[station].append(pt)
                
        for st in records_by_station:
            records_by_station[st].sort(key=lambda x: x['time'])
            
        return jsonify({'success': True, 'data': records_by_station})'''

new_stats = '''    try:
        target_id = student_info['id']
        records_by_station = {}
        bq_client, project_id = get_bq_client()
        if bq_client:
            q = f"SELECT station, timestamp, opa1_sum, opa2_sum, opa3_sum FROM \{project_id}.grading_data.grading_logs\ WHERE student_id = @sid AND (is_deleted = FALSE OR is_deleted IS NULL) ORDER BY timestamp ASC"
            job_config = bq_client.query_class.QueryJobConfig(
                query_parameters=[bq_client.query_class.ScalarQueryParameter("sid", "STRING", target_id)]
            )
            res = bq_client.query(q, job_config=job_config).result()
            
            def to_int(v):
                try: return int(v)
                except: return None
                
            for r in res:
                if r.station not in records_by_station:
                    records_by_station[r.station] = []
                records_by_station[r.station].append({
                    'time': r.timestamp.strftime('%Y/%m/%d %H:%M:%S') if r.timestamp else '',
                    'opa1': to_int(r.opa1_sum),
                    'opa2': to_int(r.opa2_sum),
                    'opa3': to_int(r.opa3_sum)
                })
                
        return jsonify({'success': True, 'data': records_by_station})'''

content = content.replace(old_stats, new_stats)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print('Done app.py stats')
