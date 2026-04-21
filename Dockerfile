# 選擇官方的輕量級 Python 映像檔 (升級至 3.11 以支援新版套件)
FROM python:3.11-slim

# 設定工作目錄
WORKDIR /app

# 將目前目錄下的檔案全部複製到容器內的 /app
COPY . /app

# 安裝相依套件與編譯工具，並安裝 Playwright 瀏覽器及其所需的系統依賴
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libffi-dev \
    libssl-dev \
    && pip install --no-cache-dir -r requirements.txt \
    # 安裝 Playwright 瀏覽器及系統套件 (針對 chromium)
    && playwright install chromium \
    && playwright install-deps chromium \
    && apt-get purge -y build-essential \
    && apt-get autoremove -y \
    && rm -rf /var/lib/apt/lists/*

# Cloud Run 會提供 PORT 環境變數，綁定 0.0.0.0
ENV PORT=8080

# 使用 gunicorn 啟動 (app.py 裡的 app 實例)
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 app:app
