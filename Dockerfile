FROM python:3.9-slim

WORKDIR /app

# 安装系统依赖、curl（用于下载中文字体）与中文字体
RUN apt-get update && apt-get install -y \
    gcc \
    curl \
    fonts-wqy-zenhei \
    fonts-wqy-microhei \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖文件
COPY requirements.txt .

# 安装Python依赖
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用文件
COPY . .

# 下载中文字体到 app/fonts，保证图中中文正常显示（不依赖系统字体）
RUN mkdir -p /app/fonts \
    && (curl -fSL -o /app/fonts/NotoSansCJKsc-Regular.otf \
        "https://raw.githubusercontent.com/googlefonts/noto-cjk/main/Sans/OTF/SimplifiedChinese/NotoSansCJKsc-Regular.otf" \
        || curl -fSL -o /app/fonts/NotoSansCJKsc-Regular.otf \
        "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/SimplifiedChinese/NotoSansCJKsc-Regular.otf") \
    || true

# 暴露端口
EXPOSE 8501

# 健康检查
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# 启动命令
ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
