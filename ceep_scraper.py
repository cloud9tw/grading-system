import asyncio
import os
from playwright.async_api import async_playwright
from dotenv import load_dotenv

load_dotenv()

import re

def clean_html(raw_html):
    """Remove HTML tags and clean up whitespace."""
    if not raw_html: return ""
    cleanr = re.compile('<.*?>')
    cleantext = re.sub(cleanr, ' ', raw_html)
    return ' '.join(cleantext.split())

async def scrape_ceep_all_forms(callback=None):
    """
    Log in to CEEP2, and scrape DOPS and Mini-CEX forms across multiple plans (Interns, PGY, New Staff).
    Returns: (final_results, task_summary)
    """
    async def report(msg):
        try:
            print(msg)
        except UnicodeEncodeError:
            print(msg.encode('utf-8', errors='replace').decode('utf-8', errors='replace'))
        
        if callback:
            await callback(msg)

    import datetime
    now = datetime.datetime.now()
    if now.month >= 7:
        N = now.year - 1911
    else:
        N = now.year - 1912
    
    # 定義抓取任務 (學年度, 計畫名稱)
    tasks = [
        (f"{N} 學年度", f"{N}學年醫事放射實習"),        # 實習學生 (目前學年)
        (f"{N} 學年度", f"醫事放射PGY {N}-影醫"),      # PGY (目前學年)
        (f"{N-1} 學年度", f"醫事放射PGY {N-1}-影醫"),    # PGY (去年)
        (f"{N} 學年度", f"影像醫學部新進放射師-{N}年"), # 新進人員 (今年)
        (f"{N-1} 學年度", f"影像醫學部新進放射師-{N-1}年") # 新進人員 (去年)
    ]

    targets = {
        "CEEP_DOPS": "醫學影像技術學-操作技能直接觀察(DOPS)評量表",
        "CEEP_MiniCEX": "醫學影像技術學-迷你臨床演練評量(Mini-CEX)評量表",
        "CEEP_TeachingRecord": "醫事放射-教學記錄"
    }
    
    final_results = {}
    task_summary = [] # 用於回傳給前端顯示

    async with async_playwright() as p:
        await report("--- 正在啟動瀏覽器核心 (Chromium) ---")
        # Launch browser with Docker-friendly arguments
        browser = await p.chromium.launch(
            headless=True,
            args=[
                '--no-sandbox', 
                '--disable-setuid-sandbox', 
                '--disable-dev-shm-usage',
                '--disable-gpu'
            ]
        )
        context = await browser.new_context(viewport={'width': 1280, 'height': 800})
        page = await context.new_page()

        await report("--- 正在連線至 CEEP 系統 ---")
        await page.goto("https://ceep2.tmu.edu.tw/")

        # Login
        await report("--- 正在執行登入作業 ---")
        await page.fill('input[name="account"]', "15680")
        await page.fill('input[name="password"]', "4249")
        
        async with page.expect_navigation():
            await page.click('button[type="submit"]')
        
        if "login" in page.url:
            await report("❌ 登入失敗：帳號或密碼錯誤")
            raise Exception("登入失敗，請確認 ceep_scraper.py 中的帳號密碼。")
        
        await report("✅ 登入成功，開始抓取流程")

        for sheet_name, form_label in targets.items():
            await report(f"➔ 準備抓取表單: {sheet_name}")
            all_records_for_form = []

            for year_label, plan_label in tasks:
                # 根據身份動態決定教學記錄表單名稱
                if sheet_name == "CEEP_TeachingRecord":
                    if "實習" in plan_label:
                        current_form_label = "醫事放射-教學記錄(實習生)"
                    elif "PGY" in plan_label:
                        current_form_label = "醫事放射-PGY教學記錄"
                    else:
                        current_form_label = "醫事放射-教學記錄" # Fallback for others
                else:
                    current_form_label = form_label

                await report(f"   [任務] {year_label} | {plan_label} -> 使用表單: {current_form_label}")
                
                # Navigate to Statistics directly
                await page.goto("https://ceep2.tmu.edu.tw/admin/complex/assessment_form/assessment_form_statistics")
                await page.wait_for_load_state("networkidle")
                
                try:
                    await page.wait_for_selector('select[name="batch_year[]"]', timeout=5000)
                    
                    # 1. 選取學年度 (手動尋找包含年度數字的選項，避免 Pattern 序列化問題)
                    year_num = year_label.split(" ")[0] # 提取 "114"
                    year_options = await page.eval_on_selector('select[name="batch_year[]"]', 
                        f"(el, yr) => Array.from(el.options).filter(o => o.text.includes(yr)).map(o => o.value)", year_num)
                    if year_options:
                        await page.select_option('select[name="batch_year[]"]', value=year_options)
                    await page.wait_for_timeout(800)
                    
                    # 2. 選取職類
                    await page.select_option('select[name="title_id"]', label="醫事放射職類")
                    await page.wait_for_timeout(1500)

                    
                    # 3. 選取計畫
                    try:
                        await page.select_option('select[name="batch_id"]', label=plan_label)
                    except:
                        # 嘗試等待連動後再選一次
                        await page.wait_for_timeout(2000)
                        await page.select_option('select[name="batch_id"]', label=plan_label)
                    await page.wait_for_timeout(1000)
                    
                    # 4. 選取表單名稱
                    await page.select_option('select[name="sf_id"]', label=current_form_label)

                    # Click Search
                    await page.click('.btn-query')
                    await page.wait_for_load_state("networkidle")
                    await page.wait_for_timeout(3000)

                    # Extract rows
                    rows = await page.query_selector_all('table.table-bordered tbody tr')
                    task_count = 0
                    for row in rows:
                        cols = await row.query_selector_all('td')
                        if len(cols) < 5: continue
                        
                        case_name = await cols[0].inner_text()
                        start_time = await cols[1].inner_text()
                        student_name = await cols[2].inner_text()
                        submit_time = await cols[3].inner_text()
                        
                        scores = {}
                        for i in range(4, len(cols)):
                            col = cols[i]
                            txt = (await col.inner_text()).strip()
                            
                            if not txt:
                                popover_content = await col.get_attribute('data-content')
                                if popover_content:
                                    txt = clean_html(popover_content)
                                else:
                                    link = await col.query_selector('a')
                                    if link:
                                        href = await link.get_attribute('href')
                                        link_text = (await link.inner_text()).strip()
                                        txt = link_text if link_text else href
                            
                            scores[f"item_{i-3}"] = txt
                        
                        all_records_for_form.append({
                            "student_name": student_name.strip(),
                            "submit_time": submit_time.strip(),
                            "case_name": case_name.strip(),
                            "start_time": f"[{plan_label}] {start_time.strip()}", # 標記計畫來源
                            "scores": scores
                        })
                        task_count += 1
                    
                    await report(f"      ✔ 抓取完畢，共 {task_count} 筆紀錄")
                    task_summary.append({
                        "form": sheet_name.replace("CEEP_", ""),
                        "year": year_label,
                        "plan": plan_label,
                        "count": task_count,
                        "status": "success"
                    })

                except Exception as e:
                    await report(f"      ⚠ 任務失敗: {plan_label} - {e}")
                    task_summary.append({
                        "form": sheet_name.replace("CEEP_", ""),
                        "year": year_label,
                        "plan": plan_label,
                        "count": 0,
                        "status": "failed",
                        "error": str(e)
                    })
                    continue

            final_results[sheet_name] = all_records_for_form
            await report(f"🎯 表單 {sheet_name} 已完成彙整")

        await report("--- 同步流程結束，正在關閉瀏覽器 ---")
        await browser.close()
        return final_results, task_summary

if __name__ == "__main__":
    import json
    data, summary = asyncio.run(scrape_ceep_all_forms())
    print("\n[Summary Overview]")
    print(json.dumps(summary, indent=2, ensure_ascii=False))