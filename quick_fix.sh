#!/bin/bash

# 腾讯云服务器快速修复脚本
# 用于解决Docker无法拉取基础镜像的问题
# 使用方法: chmod +x quick_fix.sh && ./quick_fix.sh

set -e

echo "=============================================="
echo " 腾讯云服务器 - Docker镜像拉取问题修复"
echo "=============================================="
echo ""

# 方法1: 配置Docker镜像加速器
echo "[步骤1/3] 配置Docker镜像加速器..."

# 备份原有配置
if [ -f /etc/docker/daemon.json ]; then
    sudo cp /etc/docker/daemon.json /etc/docker/daemon.json.bak.$(date +%Y%m%d_%H%M%S)
    echo "已备份原有Docker配置"
fi

# 创建新的配置
sudo mkdir -p /etc/docker
sudo tee /etc/docker/daemon.json <<-'EOF'
{
  "registry-mirrors": [
    "https://mirror.ccs.tencentyun.com",
    "https://docker.1panel.live",
    "https://hub.rat.dev",
    "https://dockerpull.com",
    "https://docker.anyhub.us.kg"
  ]
}
EOF

echo "已配置多个国内Docker镜像加速器"

# 重启Docker服务
sudo systemctl daemon-reload
sudo systemctl restart docker
echo "Docker服务已重启"
echo ""

# 方法2: 手动拉取Python基础镜像
echo "[步骤2/3] 手动拉取Python基础镜像（从国内镜像源）..."

# 尝试从多个国内源拉取
IMAGE_NAME="python:3.12-slim"
PULLED=false

# 源1: 阿里云镜像
if [ "$PULLED" = false ]; then
    echo "尝试从阿里云镜像拉取..."
    if docker pull docker.m.daocloud.io/library/$IMAGE_NAME; then
        docker tag docker.m.daocloud.io/library/$IMAGE_NAME $IMAGE_NAME
        echo "成功从阿里云镜像拉取"
        PULLED=true
    fi
fi

# 源2: 腾讯云镜像
if [ "$PULLED" = false ]; then
    echo "尝试从腾讯云镜像拉取..."
    if docker pull ccr.ccs.tencentyun.com/library/$IMAGE_NAME; then
        docker tag ccr.ccs.tencentyun.com/library/$IMAGE_NAME $IMAGE_NAME
        echo "成功从腾讯云镜像拉取"
        PULLED=true
    fi
fi

# 源3: 使用Docker镜像加速器拉取原始镜像
if [ "$PULLED" = false ]; then
    echo "尝试使用镜像加速器拉取..."
    if docker pull $IMAGE_NAME; then
        echo "成功拉取镜像"
        PULLED=true
    fi
fi

if [ "$PULLED" = true ]; then
    echo "Python基础镜像拉取成功！"
else
    echo "错误：无法从任何镜像源拉取基础镜像"
    echo "请尝试方案三（本地构建后上传）或方案四（使用腾讯云容器镜像服务）"
    exit 1
fi

echo ""

# 方法3: 构建并启动应用
echo "[步骤3/3] 构建并启动应用..."

# 修改Dockerfile使用标准镜像名（因为我们已经手动拉取了）
sed -i 's|FROM docker.m.daocloud.io/library/python:3.12-slim|FROM python:3.12-slim|' Dockerfile
sed -i 's|FROM ccr.ccs.tencentyun.com/library/python:3.12-slim|FROM python:3.12-slim|' Dockerfile

echo "开始构建Docker镜像..."
docker compose build

echo "启动服务..."
docker compose up -d

echo ""
echo "=============================================="
echo " 修复完成！"
echo "=============================================="
echo ""
echo "服务状态检查:"
docker compose ps
echo ""
echo "访问地址: http://$(curl -s ifconfig.me):5000"
echo ""
echo "查看日志: docker compose logs -f"
echo ""
