# 腾讯云服务器部署问题修复指南

## 问题描述
在腾讯云服务器上执行 `docker compose build` 时，出现以下错误：
```
failed to solve: failed to compute cache key: failed to copy: httpReadSeeker: failed open: 
could not fetch content descriptor sha256:... (application/vnd.oci.image.layer.v1.tar+gzip) 
from remote: not found
```

**根本原因**: Docker无法从Docker Hub拉取 `python:3.12-slim` 基础镜像，因为Docker Hub在国内访问受限。

---

## 解决方案（按推荐顺序）

### 方案一：使用完整的腾讯云服务器部署脚本（推荐）

此脚本会自动配置多个Docker镜像加速器，并尝试从多个国内源拉取基础镜像。

#### 步骤：

1. **上传项目到服务器**
   ```bash
   # 在本地Windows电脑执行
   scp -r d:\pptx2jpg root@YOUR_SERVER_IP:/opt/
   # 或使用WinSCP等工具上传整个项目文件夹
   ```

2. **SSH登录服务器**
   ```bash
   ssh root@YOUR_SERVER_IP
   cd /opt/pptx2jpg
   ```

3. **执行部署脚本**
   ```bash
   chmod +x tencent_deploy.sh
   ./tencent_deploy.sh
   ```

脚本会自动：
- 配置多个国内Docker镜像加速器
- 尝试从多个源（阿里云、DaoCloud、腾讯云等）拉取Python基础镜像
- 构建应用镜像
- 启动服务

---

### 方案二：本地构建后上传到服务器（最稳定）

如果方案一仍然失败，这是最可靠的方案。

#### 在本地Windows电脑执行：

1. **构建Docker镜像**
   ```powershell
   # 使用Docker Desktop
   cd d:\pptx2jpg
   docker build -t pptx2jpg:latest .
   ```

2. **导出镜像为压缩包**
   ```powershell
   docker save pptx2jpg:latest | gzip > pptx2jpg.tar.gz
   ```
   生成的文件约500MB-1GB

3. **上传到腾讯云服务器**
   
   **方法A: 使用scp命令**
   ```bash
   scp pptx2jpg.tar.gz root@YOUR_SERVER_IP:/opt/
   ```
   
   **方法B: 使用WinSCP工具**
   - 下载并安装WinSCP
   - 连接到服务器
   - 将 `pptx2jpg.tar.gz` 上传到 `/opt/` 目录

#### 在腾讯云服务器执行：

4. **SSH登录服务器**
   ```bash
   ssh root@YOUR_SERVER_IP
   ```

5. **加载Docker镜像**
   ```bash
   docker load < /opt/pptx2jpg.tar.gz
   ```

6. **拉取项目代码**
   ```bash
   cd /opt
   git clone https://github.com/zhongyingfang/Invitation-batch.git
   cd Invitation-batch
   ```

7. **修改Dockerfile**
   
   编辑 `Dockerfile`，确保第3行使用标准镜像名：
   ```dockerfile
   FROM python:3.12-slim
   ```

8. **启动服务**
   ```bash
   docker compose up -d
   ```

---

### 方案三：使用腾讯云容器镜像服务（企业推荐）

适合长期使用，速度最快。

1. **在腾讯云控制台**
   - 开通"容器镜像服务"
   - 创建个人版实例（免费）
   - 创建命名空间，例如 `my-namespace`

2. **在本地Windows电脑推送镜像**
   ```powershell
   # 登录腾讯云容器镜像服务
   docker login --username=YOUR_USERNAME ccr.ccs.tencentyun.com
   
   # 标记镜像
   docker tag pptx2jpg:latest ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest
   
   # 推送镜像
   docker push ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest
   ```

3. **在腾讯云服务器拉取**
   ```bash
   # 从腾讯云内网拉取（速度极快）
   docker pull ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest
   
   # 修改Dockerfile使用腾讯云镜像
   sed -i 's|FROM python:3.12-slim|FROM ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest|' Dockerfile
   
   # 启动服务
   docker compose up -d
   ```

---

## 故障排查

### 检查Docker镜像加速器配置
```bash
docker info | grep -A 10 "Registry Mirrors"
```

### 检查已拉取的镜像
```bash
docker images | grep python
```

### 查看构建详细日志
```bash
docker compose build --no-cache 2>&1 | tee build.log
```

### 手动测试镜像拉取
```bash
# 测试从阿里云拉取
docker pull docker.m.daocloud.io/library/python:3.12-slim

# 测试从DaoCloud拉取
docker pull docker.1panel.live/library/python:3.12-slim
```

### 检查容器状态
```bash
docker compose ps
docker compose logs -f
```

### 测试服务健康状态
```bash
curl http://localhost:5000/health
```

---

## 快速参考命令

```bash
# 停止服务
docker compose down

# 重启服务
docker compose restart

# 查看实时日志
docker compose logs -f pptx2jpg-web

# 进入容器调试
docker compose exec pptx2jpg-web bash

# 检查容器内环境
docker compose exec pptx2jpg-web python --version
docker compose exec pptx2jpg-web which soffice
docker compose exec pptx2jpg-web fc-list | head -20

# 清理无用镜像和缓存
docker system prune -f
docker builder prune -f

# 重新构建（不使用缓存）
docker compose build --no-cache
```

---

## 服务访问

部署成功后，访问：
```
http://YOUR_SERVER_IP:5000
```

**注意**: 
- 确保腾讯云服务器安全组已开放5000端口
- 如果使用域名，需要配置域名解析和防火墙规则

---

## 常见问题

### Q1: 安全组如何配置？
在腾讯云控制台：
1. 进入"云服务器" → "安全组"
2. 添加入站规则：
   - 端口: 5000
   - 协议: TCP
   - 来源: 0.0.0.0/0（或指定IP）

### Q2: 如何使用域名访问？
```bash
# 安装Nginx作为反向代理
apt update
apt install nginx -y

# 配置Nginx
cat > /etc/nginx/sites-available/pptx2jpg << 'EOF'
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://localhost:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
EOF

# 启用配置
ln -s /etc/nginx/sites-available/pptx2jpg /etc/nginx/sites-enabled/
nginx -t
systemctl restart nginx
```

### Q3: 如何处理大文件上传？
```bash
# 修改Nginx配置支持大文件
client_max_body_size 100M;

# 修改Docker Compose环境变量
environment:
  - MAX_CONTENT_LENGTH=104857600  # 100MB
```

### Q4: 如何备份数据？
```bash
# 备份上传文件和输出文件
tar -czf pptx2jpg-backup-$(date +%Y%m%d).tar.gz uploads/ web_output/

# 或使用Docker卷备份
docker compose exec pptx2jpg-web tar -czf /tmp/backup.tar.gz /app/uploads /app/web_output
docker compose cp pptx2jpg-web:/tmp/backup.tar.gz ./
```

---

## 技术支持

如遇到问题，请提供以下信息：
1. 服务器操作系统版本: `cat /etc/os-release`
2. Docker版本: `docker --version`
3. Docker Compose版本: `docker compose version`
4. 错误日志: `docker compose logs --tail=100`
5. 网络连接测试: `curl -I https://registry-1.docker.io`
