import asyncio
import sys
import os
sys.path.append(r'c:\Users\cloud\Desktop\EPA-grading\grading-system')
from ceep_scraper import scrape_ceep_all_forms

async def test():
    try:
        data, summary = await scrape_ceep_all_forms()
        print("Success!")
        print(summary)
    except Exception as e:
        print("Error occurred:")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test())
