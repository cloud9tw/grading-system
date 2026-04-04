# 選擇官方的輕量級 Python 映像檔
FROM python:3.9-slim

# 設定工作目錄
WORKDIR /app

# 將目前目錄下的檔案全部複製到容器內的 /app
COPY . /app

# 安裝相依套件
RUN pip install --no-cache-dir -r requirements.txt

# Cloud Run 會提供 PORT 環境變數，綁定 0.0.0.0
ENV PORT=8080

# 使用 gunicorn 啟動 (app.py 裡的 app 實例)
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 app:app
