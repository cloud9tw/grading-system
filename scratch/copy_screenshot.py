import shutil
import os

src = r"z:\明暉\python\debug_login_before.png"
dst = r"C:\Users\cloud\.gemini\antigravity\brain\b000ef59-ffe4-4b8c-aa1c-bf58cecca6d8\debug_login_before.png"

shutil.copy2(src, dst)
print("Copied.")
