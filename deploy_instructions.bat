@echo off
REM 腾讯云服务器部署说明
REM 在腾讯云服务器上执行以下步骤

echo ==============================================
echo   腾讯云服务器部署指南
echo ==============================================
echo.
echo 问题诊断：Docker无法拉取python:3.12-slim基础镜像
echo.
echo 解决方案（按推荐顺序）：
echo.
echo ==============================================
echo 方案一：配置Docker镜像加速器后重试（推荐）
echo ==============================================
echo.
echo 1. 上传项目文件到服务器
echo    scp -r d:\pptx2jpg root@YOUR_SERVER_IP:/opt/pptx2jpg
echo    或使用FTP/SFTP工具上传
echo.
echo 2. SSH登录到服务器
echo    ssh root@YOUR_SERVER_IP
echo.
echo 3. 执行优化部署脚本
echo    cd /opt/pptx2jpg
echo    chmod +x build_tencent.sh
echo    ./build_tencent.sh
echo.
echo 脚本会自动配置多个国内Docker镜像加速器并重新构建
echo.
echo ==============================================
echo 方案二：使用阿里云Python镜像（如果方案一失败）
echo ==============================================
echo.
echo 在服务器上手动执行：
echo.
echo 1. 编辑Dockerfile，将第1行改为：
echo    FROM docker.m.daocloud.io/library/python:3.12-slim
echo.
echo 2. 或者使用国内镜像源：
echo    FROM registry.cn-hangzhou.aliyuncs.com/library/python:3.12-slim
echo.
echo 3. 然后重新构建：
echo    docker compose build
echo    docker compose up -d
echo.
echo ==============================================
echo 方案三：本地构建后上传到服务器（最稳定）
echo ==============================================
echo.
echo 在本地电脑（Windows）执行：
echo.
echo   步骤1: 构建Docker镜像
echo   docker build -t pptx2jpg:latest .
echo.
echo   步骤2: 保存镜像为压缩包
echo   docker save pptx2jpg:latest ^| gzip ^> pptx2jpg.tar.gz
echo   （生成的文件大约500MB-1GB）
echo.
echo   步骤3: 上传到腾讯云服务器
echo   scp pptx2jpg.tar.gz root@YOUR_SERVER_IP:/opt/
echo   （或使用WinSCP等工具上传）
echo.
echo 在腾讯云服务器执行：
echo.
echo   步骤4: SSH登录服务器
echo   ssh root@YOUR_SERVER_IP
echo.
echo   步骤5: 加载镜像
echo   docker load ^< /opt/pptx2jpg.tar.gz
echo.
echo   步骤6: 拉取项目代码并启动
echo   cd /opt
echo   git clone https://github.com/zhongyingfang/Invitation-batch.git
echo   cd Invitation-batch
echo   docker compose up -d
echo.
echo ==============================================
echo 方案四：使用腾讯云容器镜像服务（企业推荐）
echo ==============================================
echo.
echo 1. 在腾讯云控制台开通"容器镜像服务"
echo 2. 创建个人版实例（免费）
echo 3. 本地推送镜像：
echo    docker tag pptx2jpg:latest ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest
echo    docker push ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest
echo 4. 服务器从内网拉取（速度极快）：
echo    docker pull ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest
echo    docker compose up -d
echo.
echo ==============================================
echo 常见问题排查
echo ==============================================
echo.
echo Q: 如何确认Docker镜像加速器是否生效？
echo A: docker info ^| grep -A 10 "Registry Mirrors"
echo.
echo Q: 如何查看构建详细日志？
echo A: docker compose build --no-cache 2^>^&1 ^| tee build.log
echo.
echo Q: 如何检查基础镜像是否已拉取？
echo A: docker images ^| grep python
echo.
echo Q: 如何进入容器调试？
echo A: docker compose exec pptx2jpg-web bash
echo.
echo ==============================================
echo 验证部署
echo ==============================================
echo.
echo 1. 检查容器状态
echo    docker compose ps
echo.
echo 2. 查看运行日志
echo    docker compose logs -f
echo.
echo 3. 访问Web界面
echo    http://YOUR_SERVER_IP:5000
echo.
echo 4. 检查健康状态
echo    curl http://localhost:5000/health
echo.
pause
