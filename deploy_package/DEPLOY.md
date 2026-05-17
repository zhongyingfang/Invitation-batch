# 批量邀请函生成器 - Docker 部署说明

## 系统要求

- **Docker**: 20.10+
- **Docker Compose**: 2.0+
- **内存**: 至少 2GB 可用内存
- **磁盘**: 至少 1GB 可用空间

## 快速部署

### 1. 解压部署包

```bash
# 解压部署包
tar -xzf pptx2jpg-docker-deploy.tar.gz
# 或
unzip pptx2jpg-docker-deploy.zip
cd pptx2jpg
```

### 2. 启动服务

```bash
# 构建并启动容器
docker compose up -d --build

# 查看日志
docker compose logs -f
```

### 3. 访问 Web 界面

打开浏览器访问：http://localhost:5000

## 使用流程

### 第一步：上传数据文件

1. 访问 http://localhost:5000
2. 上传 Excel 数据文件（`.xlsx`, `.xls`, `.xlsm`）
3. 上传模板文件（`.pptx` 或 `.docx`，至少一个）
4. 勾选 "生成 PNG 图片"（可选）
5. 点击 "开始生成"

### 第二步：等待处理完成

- 处理进度会实时显示
- 处理完成后会自动提供下载链接

### 第三步：下载结果

- 点击下载按钮获取 ZIP 压缩包
- 压缩包包含：
  - `output_documents/` - 生成的邀请函文件（DOCX/PPTX）
  - `output_images/` - 生成的 PNG 图片（如果勾选）

## Excel 数据格式

Excel 文件应包含以下列（列名必须精确匹配）：

| 列名 | 说明 | 必填 |
|------|------|------|
| 姓名 | 受邀人姓名 | ✅ |
| 单位 | 受邀人单位 | ✅ |
| 职务 | 受邀人职务（可选，不填则根据性别自动填充"先生/女士"） | ❌ |
| 性别 | 性别（可选，用于自动填充职务） | ❌ |

## 模板文件配置

### DOCX 模板

在 Word 文档中使用占位符：

```
{{姓名}} - 将被替换为姓名
{{单位}} - 将被替换为单位
{{职务}} - 将被替换为职务
```

也支持 `{姓名}`、`{单位}`、`{职务}` 格式。

### PPTX 模板

在 PowerPoint 幻灯片中使用相同的占位符格式。

## 端口配置

默认服务运行在 `5000` 端口。如需修改，编辑 `docker-compose.yml`：

```yaml
ports:
  - "8080:5000"  # 将宿主机的 8080 端口映射到容器的 5000 端口
```

## 常见问题

### Q: 生成的 PDF/PNG 断行不一致

**解决方案：**
1. 确保使用开源字体（Liberation Serif/Sans 替代 Times New Roman/Arial）
2. Docker 镜像已内置微软核心字体（Arial、Times New Roman 等）
3. 重新构建镜像：`docker compose build --no-cache`

### Q: LibreOffice 转换失败

**解决方案：**
1. 检查容器日志：`docker compose logs pptx2jpg-web`
2. 确认 LibreOffice 已安装：`docker compose exec pptx2jpg-web which soffice`
3. 确认字体已安装：`docker compose exec pptx2jpg-web fc-list | head -20`
4. 重新构建镜像：`docker compose build --no-cache && docker compose up -d`

### Q: 处理大文件超时

**解决方案：**
编辑 `docker-compose.yml` 增加超时时间：

```yaml
environment:
  - GUNICORN_TIMEOUT=600
```

### Q: 如何备份数据

```bash
# 备份上传文件
docker compose exec pptx2jpg-web tar -czf /app/uploads_backup.tar.gz /app/uploads

# 复制备份到宿主机
docker compose cp pptx2jpg-web:/app/uploads_backup.tar.gz ./
```

## 维护命令

```bash
# 停止服务
docker compose down

# 重启服务
docker compose restart

# 查看容器状态
docker compose ps

# 查看实时日志
docker compose logs -f pptx2jpg-web

# 进入容器调试
docker compose exec pptx2jpg-web bash

# 删除容器和镜像（谨慎使用）
docker compose down --rmi all --volumes
```

## 文件结构

```
pptx2jpg/
├── Dockerfile              # Docker 镜像构建文件
├── docker-compose.yml      # Docker Compose 配置
── .dockerignore           # Docker 构建忽略文件
├── requirements.txt        # Python 依赖
├── web_app.py              # Web 应用入口
├── main.py                 # 命令行入口
├── gui.py                  # GUI 入口
├── diagnose.py             # 诊断工具
├── README.md               # 项目说明
├── src/                    # 源代码
│   ├── __init__.py
│   ├── batch_processor.py  # 批量处理
│   ├── excel_reader.py     # Excel 读取
│   ├── pptx_handler.py     # PPTX 处理
│   ├── docx_handler.py     # DOCX 处理
│   ── utils.py            # 工具函数
└── templates/              # Web 模板
    └── index.html          # 主页面
```

## 技术栈

- **后端**: Flask + Gunicorn
- **文档处理**: python-docx, python-pptx
- **文件转换**: LibreOffice + pdf2image + Poppler
- **前端**: HTML + JavaScript (Server-Sent Events)

## 许可证

MIT License
