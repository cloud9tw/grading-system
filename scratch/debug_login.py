import asyncio
from playwright.async_api import async_playwright

async def debug_login():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox']
        )
        context = await browser.new_context()
        page = await context.new_page()

        print("Navigating to CEEP2...")
        await page.goto("https://ceep2.tmu.edu.tw/")
        
        print("Taking screenshot before login...")
        await page.screenshot(path="debug_login_before.png")
        
        print("Filling credentials...")
        await page.fill('input[name="account"]', "15680")
        await page.fill('input[name="password"]', "4249")
        
        print("Taking screenshot with filled credentials...")
        await page.screenshot(path="debug_login_filled.png")
        
        print("Clicking submit...")
        try:
            async with page.expect_navigation(timeout=10000):
                await page.click('button[type="submit"]')
        except Exception as e:
            print("Navigation wait failed:", e)
            
        print("Taking screenshot after submit...")
        await page.screenshot(path="debug_login_after.png")
        
        print("Current URL:", page.url)
        if "login" in page.url:
            print("Login failed. Checking for error messages...")
            error_html = await page.content()
            with open("debug_login_error.html", "w", encoding="utf-8") as f:
                f.write(error_html)
                
        await browser.close()

if __name__ == "__main__":
    asyncio.run(debug_login())
