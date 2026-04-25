import os
from privacy_utils import PrivacyManager
from dotenv import load_dotenv

load_dotenv()

def init_privacy_sheet():
    print("正在初始化隱私查照表...")
    pm = PrivacyManager()
    print("✅ 隱私查照表已準備就緒。")

if __name__ == "__main__":
    init_privacy_sheet()
