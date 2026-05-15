FROM python:3.12-slim

WORKDIR /app
ENV PYTHONUNBUFFERED=1

# USTC 镜像源 — 自动匹配基础系统的 Debian 版本代号，兼容 testing/unstable
RUN rm -rf /etc/apt/sources.list.d/* && \
    . /etc/os-release && \
    cat > /etc/apt/sources.list <<EOF
deb https://mirrors.ustc.edu.cn/debian ${VERSION_CODENAME} main contrib non-free
deb-src https://mirrors.ustc.edu.cn/debian ${VERSION_CODENAME} main contrib non-free
deb https://mirrors.ustc.edu.cn/debian ${VERSION_CODENAME}-updates main contrib non-free
deb-src https://mirrors.ustc.edu.cn/debian ${VERSION_CODENAME}-updates main contrib non-free
EOF
# testing/unstable 没有独立 -security 源，只给 stable 添加
RUN . /etc/os-release && \
    if echo "${VERSION_CODENAME}" | grep -qiE "trixi|sid|testing|unstable|next"; then \
        echo "Skipping -security repo (testing/unstable)"; \
    else \
        echo "deb https://mirrors.ustc.edu.cn/debian-security ${VERSION_CODENAME}-security main contrib non-free" >> /etc/apt/sources.list && \
        echo "deb-src https://mirrors.ustc.edu.cn/debian-security ${VERSION_CODENAME}-security main contrib non-free" >> /etc/apt/sources.list; \
    fi

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-writer \
       libreoffice-impress \
       poppler-utils \
       fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -i https://mirrors.tencentyun.com/pypi/simple -r requirements.txt

COPY . ./

# 创建数据目录
RUN mkdir -p uploads web_output

EXPOSE 5000
HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:5000/health')" || exit 1

CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "600", "--graceful-timeout", "30", "web_app:app"]
