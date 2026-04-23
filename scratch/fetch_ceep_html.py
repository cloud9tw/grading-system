import urllib.request

try:
    req = urllib.request.Request("https://ceep2.tmu.edu.tw/", headers={'User-Agent': 'Mozilla/5.0'})
    html = urllib.request.urlopen(req).read().decode('utf-8')
    with open("debug_ceep_home.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("HTML saved.")
except Exception as e:
    print("Error:", e)
