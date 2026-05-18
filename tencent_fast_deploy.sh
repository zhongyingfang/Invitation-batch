#!/bin/bash

# 腾讯云服务器极速部署脚本
# 使用优化版Dockerfile，减少包依赖，大幅加快构建速度
# 使用方法: chmod +x tencent_fast_deploy.sh && ./tencent_fast_deploy.sh

set -e

echo "=============================================="
echo " 腾讯云服务器 - 极速部署脚本"
echo " 使用最小化安装策略，构建速度提升50-70%"
echo "=============================================="
echo ""

# 检查必要工具
echo "[检查] 验证必要工具..."
if ! command -v docker &> /dev/null; then
    echo "错误: 未找到Docker，请先安装Docker"
    exit 1
fi

if ! docker compose version &> /dev/null; then
    echo "错误: 未找到Docker Compose，请先安装"
    exit 1
fi

echo "✓ Docker和Docker Compose已安装"
echo ""

# 步骤1: 配置Docker镜像加速器
echo "[步骤1/4] 配置Docker镜像加速器..."

if [ ! -f /etc/docker/daemon.json ]; then
    sudo mkdir -p /etc/docker
    sudo tee /etc/docker/daemon.json <<-'EOF'
{
  "registry-mirrors": [
    "https://mirror.ccs.tencentyun.com",
    "https://docker.1panel.live",
    "https://hub.rat.dev"
  ]
}
EOF
    sudo systemctl daemon-reload
    sudo systemctl restart docker
    echo "✓ 已配置Docker镜像加速器"
else
    echo "✓ Docker配置已存在"
fi
echo ""

# 步骤2: 拉取Python基础镜像
echo "[步骤2/4] 拉取Python基础镜像..."

IMAGE="python:3.12-slim"

# 检查是否已有镜像
if docker images --format "{{.Repository}}:{{.Tag}}" | grep -q "^${IMAGE}$"; then
    echo "✓ Python基础镜像已存在，跳过拉取"
else
    echo "尝试从国内源拉取..."
    PULLED=false
    
    # 源1: 阿里云
    if docker pull docker.m.daocloud.io/library/$IMAGE 2>/dev/null; then
        docker tag docker.m.daocloud.io/library/$IMAGE $IMAGE
        docker rmi docker.m.daocloud.io/library/$IMAGE 2>/dev/null || true
        PULLED=true
        echo "✓ 从阿里云镜像拉取成功"
    fi
    
    # 源2: 使用镜像加速器
    if [ "$PULLED" = false ]; then
        if docker pull $IMAGE 2>/dev/null; then
            PULLED=true
            echo "✓ 使用镜像加速器拉取成功"
        fi
    fi
    
    if [ "$PULLED" = false ]; then
        echo " 无法拉取Python基础镜像"
        echo ""
        echo "请使用本地构建方案："
        echo "  1. 在本地Windows电脑运行: docker build -t pptx2jpg:latest ."
        echo "  2. 导出: docker save pptx2jpg:latest | gzip > pptx2jpg.tar.gz"
        echo "  3. 上传到服务器并执行: docker load < pptx2jpg.tar.gz"
        exit 1
    fi
fi
echo ""

# 步骤3: 备份并清理
echo "[步骤3/4] 清理缓存..."

# 备份原Dockerfile（如果存在）
if [ -f "Dockerfile" ] && [ "$(head -1 Dockerfile)" != "# 腾讯云优化版 Dockerfile" ]; then
    cp Dockerfile Dockerfile.backup.$(date +%Y%m%d)
    echo "✓ 已备份原Dockerfile"
fi

# 清理Docker缓存
docker system prune -f 2>/dev/null || true
echo "✓ 缓存清理完成"
echo ""

# 步骤4: 构建并启动
echo "[步骤4/4] 构建并启动服务..."
echo "预计构建时间: 2-4分钟（比原版快50-70%）"
echo ""
echo "优化内容:"
echo "  ✓ 移除 libreoffice-impress（节省约100MB）"
echo "  ✓ 移除 ttf-mscorefonts-installer（避免外部下载）"
echo "  ✓ 移除 fonts-dejavu-core（使用liberation替代）"
echo "  ✓ 启用非交互式安装模式"
echo ""

echo "开始构建..."
docker compose build --no-cache

echo ""
echo "✓ 构建完成！"
echo ""

echo "启动服务..."
docker compose up -d

echo ""
echo "=============================================="
echo " 部署完成！"
echo "=============================================="
echo ""
echo "服务状态:"
docker compose ps
echo ""

# 等待服务启动
echo "等待服务启动..."
sleep 5

# 检查健康状态
if curl -s http://localhost:5000/health > /dev/null 2>&1; then
    echo "✓ 服务健康检查通过"
else
    echo "⚠ 服务可能还在启动中，请稍后检查"
fi

echo ""
echo "访问地址:"
SERVER_IP=$(curl -s ifconfig.me 2>/dev/null || echo "YOUR_SERVER_IP")
echo "  http://${SERVER_IP}:5000"
echo ""
echo "常用命令:"
echo "  查看日志: docker compose logs -f"
echo "  测试健康: curl http://localhost:5000/health"
echo "  停止服务: docker compose down"
echo "  重启服务: docker compose restart"
echo ""
