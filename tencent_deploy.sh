#!/bin/bash

# 腾讯云服务器完整部署脚本
# 使用方法: chmod +x tencent_deploy.sh && ./tencent_deploy.sh

set -e

echo "=============================================="
echo " 腾讯云服务器 - 完整部署脚本"
echo "=============================================="
echo ""

# 检查是否安装了必要工具
echo "[检查] 验证必要工具..."
if ! command -v docker &> /dev/null; then
    echo "错误: 未找到Docker，请先安装Docker"
    exit 1
fi

if ! command -v docker-compose &> /dev/null && ! docker compose version &> /dev/null; then
    echo "错误: 未找到Docker Compose，请先安装"
    exit 1
fi

echo "✓ Docker和Docker Compose已安装"
echo ""

# 步骤1: 配置Docker镜像加速器
echo "[步骤1/5] 配置Docker镜像加速器..."

BACKUP_DIR="/etc/docker/backups"
sudo mkdir -p $BACKUP_DIR

if [ -f /etc/docker/daemon.json ]; then
    sudo cp /etc/docker/daemon.json $BACKUP_DIR/daemon.json.bak.$(date +%Y%m%d_%H%M%S)
    echo "✓ 已备份原有Docker配置"
fi

sudo tee /etc/docker/daemon.json <<-'EOF'
{
  "registry-mirrors": [
    "https://mirror.ccs.tencentyun.com",
    "https://docker.1panel.live",
    "https://hub.rat.dev",
    "https://dockerpull.com",
    "https://docker.anyhub.us.kg",
    "https://dockerproxy.com"
  ]
}
EOF

echo "✓ 已配置多个国内Docker镜像加速器"

sudo systemctl daemon-reload
sudo systemctl restart docker
echo "✓ Docker服务已重启"
echo ""

# 步骤2: 尝试拉取Python基础镜像
echo "[步骤2/5] 拉取Python基础镜像..."

IMAGE="python:3.12-slim"
PULLED=false

echo "尝试从多个国内源拉取 $IMAGE ..."
echo ""

# 源1: 阿里云
if [ "$PULLED" = false ]; then
    echo "  → 尝试阿里云镜像 (docker.m.daocloud.io)..."
    if docker pull docker.m.daocloud.io/library/$IMAGE 2>/dev/null; then
        docker tag docker.m.daocloud.io/library/$IMAGE $IMAGE
        docker rmi docker.m.daocloud.io/library/$IMAGE 2>/dev/null || true
        echo "  ✓ 成功从阿里云镜像拉取"
        PULLED=true
    fi
fi

# 源2: DaoCloud
if [ "$PULLED" = false ]; then
    echo "  → 尝试DaoCloud镜像 (docker.m.daocloud.io)..."
    if docker pull docker.m.daocloud.io/library/$IMAGE 2>/dev/null; then
        docker tag docker.m.daocloud.io/library/$IMAGE $IMAGE
        docker rmi docker.m.daocloud.io/library/$IMAGE 2>/dev/null || true
        echo "  ✓ 成功从DaoCloud镜像拉取"
        PULLED=true
    fi
fi

# 源3: 使用配置好的镜像加速器拉取
if [ "$PULLED" = false ]; then
    echo "  → 尝试使用镜像加速器拉取..."
    if docker pull $IMAGE 2>/dev/null; then
        echo "  ✓ 成功拉取镜像"
        PULLED=true
    fi
fi

# 源4: 尝试1panel镜像
if [ "$PULLED" = false ]; then
    echo "  → 尝试1Panel镜像 (docker.1panel.live)..."
    if docker pull docker.1panel.live/library/$IMAGE 2>/dev/null; then
        docker tag docker.1panel.live/library/$IMAGE $IMAGE
        docker rmi docker.1panel.live/library/$IMAGE 2>/dev/null || true
        echo "  ✓ 成功从1Panel镜像拉取"
        PULLED=true
    fi
fi

if [ "$PULLED" = true ]; then
    echo ""
    echo "✓ Python基础镜像拉取成功！"
    docker images $IMAGE
else
    echo ""
    echo " 错误：无法从任何镜像源拉取基础镜像"
    echo ""
    echo "请尝试以下解决方案："
    echo ""
    echo "方案A: 本地构建后上传（推荐）"
    echo "  1. 在本地电脑运行: docker build -t pptx2jpg:latest ."
    echo "  2. 导出镜像: docker save pptx2jpg:latest | gzip > pptx2jpg.tar.gz"
    echo "  3. 上传到服务器: scp pptx2jpg.tar.gz root@服务器IP:/"
    echo "  4. 在服务器加载: docker load < pptx2jpg.tar.gz"
    echo ""
    echo "方案B: 手动修改Dockerfile"
    echo "  编辑Dockerfile，将第1行改为："
    echo "  FROM docker.1panel.live/library/python:3.12-slim"
    echo ""
    echo "方案C: 配置代理"
    echo "  export http_proxy=http://代理地址:端口"
    echo "  export https_proxy=http://代理地址:端口"
    echo "  然后重新运行此脚本"
    echo ""
    exit 1
fi

echo ""

# 步骤3: 检查项目文件
echo "[步骤3/5] 检查项目文件..."

if [ ! -f "Dockerfile" ]; then
    echo "✗ 错误: 未找到Dockerfile，请确保在项目根目录执行此脚本"
    exit 1
fi

if [ ! -f "docker-compose.yml" ]; then
    echo "✗ 错误: 未找到docker-compose.yml"
    exit 1
fi

echo "✓ 项目文件检查通过"
echo ""

# 步骤4: 构建应用镜像
echo "[步骤4/5] 构建应用Docker镜像..."

# 修改Dockerfile使用标准镜像名
if grep -q "docker.m.daocloud.io" Dockerfile; then
    sed -i 's|FROM docker.m.daocloud.io/library/python:3.12-slim|FROM python:3.12-slim|' Dockerfile
    echo "✓ 已修改Dockerfile使用标准镜像名"
fi

if grep -q "ccr.ccs.tencentyun.com" Dockerfile; then
    sed -i 's|FROM ccr.ccs.tencentyun.com/library/python:3.12-slim|FROM python:3.12-slim|' Dockerfile
    echo "✓ 已修改Dockerfile使用标准镜像名"
fi

echo "开始构建..."
docker compose build --no-cache

echo "✓ 应用镜像构建成功！"
echo ""

# 步骤5: 启动服务
echo "[步骤5/5] 启动服务..."

docker compose up -d

echo ""
echo "=============================================="
echo " 部署完成！"
echo "=============================================="
echo ""
echo "服务状态:"
docker compose ps
echo ""
echo "服务访问地址:"
echo "  http://$(curl -s ifconfig.me 2>/dev/null || echo 'YOUR_SERVER_IP'):5000"
echo ""
echo "常用命令:"
echo "  查看日志: docker compose logs -f"
echo "  停止服务: docker compose down"
echo "  重启服务: docker compose restart"
echo "  进入容器: docker compose exec pptx2jpg-web bash"
echo "  检查健康: curl http://localhost:5000/health"
echo ""
