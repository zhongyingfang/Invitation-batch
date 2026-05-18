@echo off
echo ==============================================
echo   本地构建Docker镜像并导出
echo ==============================================
echo.
echo 此脚本将在本地构建Docker镜像并导出为压缩包
echo 然后您可以上传到腾讯云服务器
echo.
pause

echo.
echo [1/3] 清理旧的构建缓存...
docker system prune -f

echo.
echo [2/3] 构建Docker镜像...
docker build -t pptx2jpg:latest .

if %errorlevel% neq 0 (
    echo.
    echo 错误：Docker镜像构建失败！
    echo 请检查Docker是否正在运行，以及网络连接是否正常
    pause
    exit /b 1
)

echo.
echo [3/3] 导出镜像为压缩包...
docker save pptx2jpg:latest | gzip > pptx2jpg.tar.gz

if %errorlevel% neq 0 (
    echo.
    echo 错误：镜像导出失败！
    pause
    exit /b 1
)

echo.
echo ==============================================
echo   构建完成！
echo ==============================================
echo.
echo 生成的文件: pptx2jpg.tar.gz
echo 文件大小: 大约 500MB - 1GB
echo.
echo 下一步：
echo 1. 将 pptx2jpg.tar.gz 上传到腾讯云服务器
echo    使用WinSCP、FileZilla或scp命令
echo.
echo 2. 在服务器上执行：
echo    docker load ^< pptx2jpg.tar.gz
echo    cd /path/to/project
echo    docker compose up -d
echo.
echo 查看生成的文件：
dir pptx2jpg.tar.gz
echo.
pause
