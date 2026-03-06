# 選擇 Python 3.11 slim 版本
FROM python:3.11-slim

# 安裝系統套件：Graphviz 與中文字型
RUN apt-get update && \
    apt-get install -y graphviz fonts-noto-cjk && \
    rm -rf /var/lib/apt/lists/*

# 設定工作目錄
WORKDIR /app

# 複製需求檔並安裝 Python 套件
COPY requirements.txt .
RUN pip install --upgrade pip setuptools wheel
RUN pip install --no-cache-dir -r requirements.txt

# 複製整個專案程式碼
COPY . .

# Web Service 啟動命令
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:10000"]