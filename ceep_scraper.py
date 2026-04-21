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

async def scrape_ceep_all_forms():
    """
    Log in to CEEP2, and scrape DOPS and Mini-CEX forms across multiple plans (Interns, PGY, New Staff).
    Returns: (final_results, task_summary)
    """
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
        (f"{N} 學年度", f"影像醫學部新進放射師-{N}年"), # 新進人員 (今年)
        (f"{N-1} 學年度", f"影像醫學部新進放射師-{N-1}年") # 新進人員 (去年)
    ]

    targets = {
        "CEEP_DOPS": "醫學影像技術學-操作技能直接觀察(DOPS)評量表",
        "CEEP_MiniCEX": "醫學影像技術學-迷你臨床演練評量(Mini-CEX)評量表"
    }
    
    final_results = {}
    task_summary = [] # 用於回傳給前端顯示

    async with async_playwright() as p:
        # Launch browser
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(viewport={'width': 1280, 'height': 800})
        page = await context.new_page()

        print("--- 正在連線至 CEEP 系統 ---")
        await page.goto("https://ceep2.tmu.edu.tw/")

        # Login
        await page.fill('input[name="account"]', "15680")
        await page.fill('input[name="password"]', "4249")
        
        async with page.expect_navigation():
            await page.click('button[type="submit"]')
        
        if "login" in page.url:
            raise Exception("登入失敗，請確認 ceep_scraper.py 中的帳號密碼。")

        for sheet_name, form_label in targets.items():
            print(f"\n===== 開始抓取表單: {form_label} =====")
            all_records_for_form = []

            for year_label, plan_label in tasks:
                print(f"--- 執行任務: [{year_label}] {plan_label} ---")
                
                # Navigate to Statistics directly
                await page.goto("https://ceep2.tmu.edu.tw/admin/complex/assessment_form/assessment_form_statistics")
                await page.wait_for_load_state("networkidle")
                
                try:
                    await page.wait_for_selector('select[name="batch_year[]"]', timeout=5000)
                    
                    # 1. 選取學年度
                    await page.select_option('select[name="batch_year[]"]', label=year_label)
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
                    await page.select_option('select[name="sf_id"]', label=form_label)

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
                    
                    print(f"   -> 成功抓取 {task_count} 筆紀錄")
                    task_summary.append({
                        "form": "DOPS" if "DOPS" in sheet_name else "Mini-CEX",
                        "year": year_label,
                        "plan": plan_label,
                        "count": task_count,
                        "status": "success"
                    })

                except Exception as e:
                    print(f"   ⚠️ 任務跳過或失敗: {plan_label} - {e}")
                    task_summary.append({
                        "form": "DOPS" if "DOPS" in sheet_name else "Mini-CEX",
                        "year": year_label,
                        "plan": plan_label,
                        "count": 0,
                        "status": "failed",
                        "error": str(e)
                    })
                    continue

            final_results[sheet_name] = all_records_for_form
            print(f"✅ {sheet_name} 總計完工，全計畫累計 {len(all_records_for_form)} 筆紀錄")

        await browser.close()
        return final_results, task_summary

if __name__ == "__main__":
    import json
    data, summary = asyncio.run(scrape_ceep_all_forms())
    print("\n[Summary Overview]")
    print(json.dumps(summary, indent=2, ensure_ascii=False))
