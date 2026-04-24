import shutil
import os

src = r"C:\Users\cloud\.gemini\antigravity\brain\b000ef59-ffe4-4b8c-aa1c-bf58cecca6d8\login_bg_premium_1776992876565.png"
dst = r"c:\Users\cloud\Desktop\EPA-grading\grading-system\static\login_bg.png"

# Ensure static dir exists
os.makedirs(os.path.dirname(dst), exist_ok=True)

shutil.copy2(src, dst)
print("Copied successfully.")
