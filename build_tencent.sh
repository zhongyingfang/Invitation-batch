#!/bin/bash

# 腾讯云服务器优化构建脚本
# 使用方法: chmod +x build_tencent.sh && ./build_tencent.sh

set -e

echo "=============================================="
echo " 腾讯云服务器 - Docker 镜像构建优化脚本"
echo "=============================================="

# 配置Docker使用国内镜像加速器
echo "[1/4] 配置Docker镜像加速器..."
if [ ! -f /etc/docker/daemon.json ]; then
    sudo mkdir -p /etc/docker
    sudo tee /etc/docker/daemon.json <<-'EOF'
{
  "registry-mirrors": [
    "https://mirror.ccs.tencentyun.com",
    "https://docker.1panel.live",
    "https://hub.rat.dev",
    "https://dockerpull.com"
  ]
}
EOF
    echo "已配置多个Docker镜像加速器（腾讯云+其他国内源）"
    sudo systemctl daemon-reload
    sudo systemctl restart docker
    echo "Docker服务已重启"
else
    echo "检测到已有Docker配置，备份并更新..."
    sudo cp /etc/docker/daemon.json /etc/docker/daemon.json.bak
    
    # 合并配置，添加更多镜像源
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
    echo "已更新Docker镜像加速器配置"
    sudo systemctl daemon-reload
    sudo systemctl restart docker
    echo "Docker服务已重启"
fi

# 清理旧的构建缓存
echo ""
echo "[2/4] 清理Docker构建缓存..."
docker system prune -f
docker builder prune -f

# 构建镜像（使用腾讯云镜像源）
echo ""
echo "[3/4] 开始构建Docker镜像..."
echo "这可能需要几分钟时间，请耐心等待..."
docker compose build --no-cache

# 启动服务
echo ""
echo "[4/4] 启动服务..."
docker compose up -d

echo ""
echo "=============================================="
echo " 部署完成！"
echo "=============================================="
echo ""
echo "服务访问地址: http://$(curl -s ifconfig.me):5000"
echo ""
echo "常用命令:"
echo "  查看日志: docker compose logs -f"
echo "  停止服务: docker compose down"
echo "  重启服务: docker compose restart"
echo "  进入容器: docker compose exec pptx2jpg-web bash"
echo ""
