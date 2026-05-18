@echo off
REM 腾讯云服务器部署说明
REM 在腾讯云服务器上执行以下步骤

echo ==============================================
echo   腾讯云服务器部署指南
echo ==============================================
echo.
echo 请按以下步骤在腾讯云服务器上操作：
echo.
echo 1. 上传项目文件到服务器
echo    scp -r d:\pptx2jpg root@YOUR_SERVER_IP:/opt/pptx2jpg
echo    或使用FTP/SFTP工具上传
echo.
echo 2. SSH登录到服务器
echo    ssh root@YOUR_SERVER_IP
echo.
echo 3. 进入项目目录
echo    cd /opt/pptx2jpg
echo.
echo 4. 执行优化构建脚本
echo    chmod +x build_tencent.sh
echo    ./build_tencent.sh
echo.
echo 5. 访问Web界面
echo    http://YOUR_SERVER_IP:5000
echo.
echo ==============================================
echo   如果构建仍然很慢，使用预构建镜像方案
echo ==============================================
echo.
echo 方案A: 本地构建后推送镜像
echo.
echo   1. 在本地电脑构建镜像:
echo      docker build -t pptx2jpg:latest .
echo.
echo   2. 标记并推送到腾讯云容器镜像服务:
echo      docker tag pptx2jpg:latest ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest
echo      docker push ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest
echo.
echo   3. 在服务器上拉取并运行:
echo      docker pull ccr.ccs.tencentyun.com/YOUR_NAMESPACE/pptx2jpg:latest
echo      docker compose up -d
echo.
echo 方案B: 使用离线传输方式
echo.
echo   1. 在本地电脑保存镜像:
echo      docker save pptx2jpg:latest ^| gzip ^> pptx2jpg.tar.gz
echo.
echo   2. 传输到服务器:
echo      scp pptx2jpg.tar.gz root@YOUR_SERVER_IP:/opt/
echo.
echo   3. 在服务器加载镜像:
echo      docker load ^< /opt/pptx2jpg.tar.gz
echo      docker compose up -d
echo.
pause
