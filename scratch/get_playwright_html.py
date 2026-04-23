import asyncio
from playwright.async_api import async_playwright

async def get_html():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox']
        )
        context = await browser.new_context()
        page = await context.new_page()

        print("Navigating to CEEP2...")
        try:
            await page.goto("https://ceep2.tmu.edu.tw/", wait_until="networkidle")
            html = await page.content()
            with open(r"c:\Users\cloud\Desktop\EPA-grading\grading-system\scratch\ceep_playwright.html", "w", encoding="utf-8") as f:
                f.write(html)
            print("HTML saved.")
        except Exception as e:
            print("Error:", e)
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(get_html())
