#!/bin/bash
# Tencent Cloud 部署脚本
set -e

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$APP_DIR"

echo "========================================"
echo "  PPTX2JPG - Tencent Cloud 部署"
echo "========================================"

# 检查 Docker
if ! command -v docker &>/dev/null; then
    echo "正在安装 Docker..."
    curl -fsSL https://mirrors.tencentyun.com/docker-ce/linux/debian/gpg | apt-key add -
    echo "deb https://mirrors.tencentyun.com/docker-ce/linux/debian bookworm stable" > /etc/apt/sources.list.d/docker.list
    apt-get update && apt-get install -y docker-ce docker-ce-cli containerd.io docker-compose-plugin
    systemctl enable --now docker
fi

# 创建数据目录
mkdir -p uploads web_output

# 构建并启动
echo "构建 Docker 镜像..."
docker compose build --no-cache

echo "启动服务..."
docker compose up -d

echo ""
echo "========================================"
echo "  部署完成！"
echo "  访问地址: http://$(curl -s http://metadata.tencentyun.com/latest/meta-data/public-ipv4):5000"
echo "========================================"
