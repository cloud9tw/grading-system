import asyncio
import os
from playwright.async_api import async_playwright
from dotenv import load_dotenv

load_dotenv()

async def scrape_ceep_all_forms():
    """
    Log in to CEEP2, and scrape both DOPS and Mini-CEX forms.
    Returns: { "CEEP_DOPS": [...], "CEEP_MiniCEX": [...] }
    """
    targets = {
        "CEEP_DOPS": "醫學影像技術學-操作技能直接觀察(DOPS)評量表",
        "CEEP_MiniCEX": "醫學影像技術學-迷你臨床演練評量(Mini-CEX)評量表"
    }
    
    final_results = {}

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
            print(f"--- 正在準備抓取表單: {form_label} ---")
            
            # Navigate to Statistics directly
            await page.goto("https://ceep2.tmu.edu.tw/admin/complex/assessment_form/assessment_form_statistics")
            await page.wait_for_load_state("networkidle")
            
            await page.wait_for_selector('select[name="batch_year[]"]', timeout=15000)

            # 1. 學年度: 114 學年度
            await page.select_option('select[name="batch_year[]"]', label="114 學年度")
            await page.wait_for_timeout(800)
            
            # 2. 職類: 醫事放射職類
            await page.select_option('select[name="title_id"]', label="醫事放射職類")
            await page.wait_for_timeout(1500)
            
            # 3. 表單所屬計畫: 114學年醫事放射實習
            try:
                await page.select_option('select[name="batch_id"]', label="114學年醫事放射實習")
            except:
                await page.wait_for_timeout(2000)
                await page.select_option('select[name="batch_id"]', label="114學年醫事放射實習")
            await page.wait_for_timeout(1500)
            
            # 4. 表單名稱: 依照目標表單選擇
            await page.select_option('select[name="sf_id"]', label=form_label)

            # Click Search
            print(f"--- 點擊查詢: {sheet_name} ---")
            await page.click('.btn-query')
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(3000)

            # Extract rows
            rows = await page.query_selector_all('table.table-bordered tbody tr')
            
            form_records = []
            for row in rows:
                cols = await row.query_selector_all('td')
                if len(cols) < 5: continue
                
                case_name = await cols[0].inner_text()
                start_time = await cols[1].inner_text()
                student_name = await cols[2].inner_text()
                submit_time = await cols[3].inner_text()
                
                scores = {}
                for i in range(4, len(cols)):
                    val = await cols[i].inner_text()
                    scores[f"item_{i-3}"] = val.strip()
                
                form_records.append({
                    "student_name": student_name.strip(),
                    "submit_time": submit_time.strip(),
                    "case_name": case_name.strip(),
                    "start_time": start_time.strip(),
                    "scores": scores
                })

            final_results[sheet_name] = form_records
            print(f"✅ {sheet_name} 抓取完成，共 {len(form_records)} 筆紀錄")

        await browser.close()
        return final_results

if __name__ == "__main__":
    import json
    data = asyncio.run(scrape_ceep_all_forms())
    for sheet, records in data.items():
        print(f"\n[{sheet}] First record sample:")
        if records:
             print(json.dumps(records[0], indent=2, ensure_ascii=False))
        else:
             print("No records found.")
