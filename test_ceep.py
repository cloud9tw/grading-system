import requests
from bs4 import BeautifulSoup
import os

def test_ceep_login():
    login_url = "https://ceep2.tmu.edu.tw/Login.aspx" # 假設的登入路徑，通常是 Login.aspx 或 login
    # 根據之前的觀察，嘗試模擬登入
    session = requests.Session()
    
    try:
        # 第一步：獲取登入頁面的 ViewState 等隱藏欄位 (ASP.NET 常用)
        response = session.get(login_url, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        viewstate = soup.find('input', {'id': '__VIEWSTATE'}).get('value') if soup.find('input', {'id': '__VIEWSTATE'}) else ''
        eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'}).get('value') if soup.find('input', {'id': '__EVENTVALIDATION'}) else ''
        
        # 準備登入資料 (這是基於 ASP.NET 頁面的猜測，稍後會根據執行結果調整)
        payload = {
            '__VIEWSTATE': viewstate,
            '__EVENTVALIDATION': eventvalidation,
            'ctl00$ContentPlaceHolder1$txtAccount': '15680',
            'ctl00$ContentPlaceHolder1$txtPassword': '4249',
            'ctl00$ContentPlaceHolder1$btnLogin': '登入'
        }
        
        post_response = session.post(login_url, data=payload, timeout=10)
        
        print(f"Status Code: {post_response.status_code}")
        print(f"Final URL: {post_response.url}")
        
        # 檢查是否登入成功 (判斷頁面是否有登出按鈕或教師姓名)
        if "15680" in post_response.text or "登出" in post_response.text:
            print("✅ Login Successful!")
            # 存下 HTML 以供分析結構
            with open('ceep_dashboard.html', 'w', encoding='utf-8') as f:
                f.write(post_response.text)
        else:
            print("❌ Login Failed. Check payload or HTML structure.")
            with open('login_failed.html', 'w', encoding='utf-8') as f:
                f.write(post_response.text)
                
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    test_ceep_login()
