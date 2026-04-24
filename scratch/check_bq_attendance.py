import os
from credentials_utils import get_bq_client
import datetime
try:
    client, project_id = get_bq_client()
    today_str = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime('%Y-%m-%d')
    query = f"SELECT * FROM `{project_id}.grading_data.attendance_events` WHERE event_time >= '{today_str}'"
    results = client.query(query).result()
    rows = [dict(row) for row in results]
    print(f'BQ Today Events: {len(rows)} found')
    for r in rows:
        print(r)
except Exception as e:
    print(f'BQ Error: {e}')
