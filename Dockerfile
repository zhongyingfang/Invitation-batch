FROM python:3.12-slim

WORKDIR /app
ENV PYTHONUNBUFFERED=1

RUN cat > /etc/apt/sources.list <<'EOF'
deb https://mirrors.ustc.edu.cn/debian bookworm main contrib non-free
deb-src https://mirrors.ustc.edu.cn/debian bookworm main contrib non-free
deb https://mirrors.ustc.edu.cn/debian-security bookworm-security main contrib non-free
deb-src https://mirrors.ustc.edu.cn/debian-security bookworm-security main contrib non-free
deb https://mirrors.ustc.edu.cn/debian bookworm-updates main contrib non-free
deb-src https://mirrors.ustc.edu.cn/debian bookworm-updates main contrib non-free
EOF

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-writer \
       libreoffice-impress \
       poppler-utils \
       fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt

COPY . ./

EXPOSE 5000
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "600", "--graceful-timeout", "30", "web_app:app"]
