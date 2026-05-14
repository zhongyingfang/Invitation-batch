# PPTX 和 DOCX 批量邀请函生成系统

这是一个自动化批量生成邀请函的系统，用于从Excel表格导入数据并填充到PPTX和DOCX模板文件中，最后转换成高清PNG图片。

## 项目结构

```
pptx2jpg/
├── 中国赣州家具博览会邀请表.xlsx       # 邀请嘉宾数据源
├── 电子邀请函模板.pptx               # PPTX邀请函模板
├── （人民政府）邀请函模板.docx       # DOCX邀请函模板
├── src/                              # 源代码目录
│   ├── excel_reader.py              # Excel数据读取模块
│   ├── pptx_handler.py              # PPTX处理模块
│   ├── docx_handler.py              # DOCX处理模块
│   └── batch_processor.py           # 批量处理主程序
├── output_documents/                 # 生成的PPTX/DOCX文件输出目录
├── output_images/                    # 生成的PNG图片输出目录
├── requirements.txt                  # Python依赖列表
└── README.md                         # 本说明文件
```

## 主要功能

1. **Excel数据读取**: 从Excel表格读取邀请嘉宾信息（姓名、单位、职务等）
2. **PPTX模板填充**: 使用Excel数据填充PPTX邀请函模板
3. **DOCX模板填充**: 使用Excel数据填充DOCX邀请函模板
4. **PNG转换**: 将生成的PPTX和DOCX文件转换为高清PNG图片（便于打印）

## 使用说明

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 准备模板文件

在Excel表格中：
- 第2行包含列标题（姓名、单位、职务等）
- 第3行开始为数据行

在PPTX/DOCX模板中：
- 使用占位符标记需要填充的位置，支持格式：
  - `{{姓名}}` 或 `{姓名}`
  - `{{name}}` 或 `{name}`
  - `{{单位}}` 或 `{单位}`
  - `{{unit}}` 或 `{unit}`
  - `{{职务}}` 或 `{职务}`
  - `{{position}}` 或 `{position}`

### 3. 运行程序

```bash
cd src
python batch_processor.py
```

### 4. Windows 图形界面

在项目根目录下运行：

```bash
python gui.py
```

该界面允许你添加多个 Excel + 模板映射，每个映射包含：
- 一个 Excel 数据文件
- 一个 PPTX 模板文件
- 一个 DOCX 模板文件

你可以重复添加映射，批量处理多个 Excel 与多个模板的组合。

### 5. Docker Web 版本

项目新增了 Docker Web 界面版本，无需依赖客户本机资源，所有文件处理和 PNG 转换都在容器内完成。

构建镜像：

```bash
docker build -t pptx2jpg-web .
```

运行容器：

```bash
docker run --rm -p 5000:5000 pptx2jpg-web
```

打开浏览器访问：

```text
http://localhost:5000
```

在网页界面中上传：
- Excel 数据文件
- PPTX 模板文件
- DOCX 模板文件

生成完成后会直接下载 ZIP 包，包含生成的文档和 PNG 图片。

如果使用 Docker Compose：

```bash
docker compose up --build
```

## 输出说明

- **output_documents/**: 包含所有生成的PPTX和DOCX文件
  - `电子邀请函_[姓名].pptx`
  - `政府邀请函_[姓名].docx`

- **output_images/**: 包含转换后的高清PNG图片
  - `电子邀请函/` - 电子邀请函转换后的PNG
  - `政府邀请函/` - 政府邀请函转换后的PNG

## 系统要求

- Python 3.7+
- LibreOffice（用于转换PPTX/DOCX到PDF再到PNG）

### Linux系统安装LibreOffice:
```bash
sudo apt-get install libreoffice
```

### MacOS系统:
```bash
brew install libreoffice
```

### Windows系统:
从 https://www.libreoffice.org 下载安装

## 模块说明

### excel_reader.py
- `ExcelReader` 类: 负责读取Excel文件并提取数据

### pptx_handler.py
- `PPTXHandler` 类: 
  - `fill_template()`: 填充PPTX模板
  - `convert_to_png()`: 将PPTX转换为PNG

### docx_handler.py
- `DOCXHandler` 类:
  - `fill_template()`: 填充DOCX模板
  - `convert_to_png()`: 将DOCX转换为PNG

### batch_processor.py
- `BatchProcessor` 类: 协调所有模块进行批量处理
- `main()`: 程序入口

## 常见问题

Q: 如何自定义占位符？
A: 在处理器中修改占位符识别逻辑，或使用标准格式 `{{字段名}}` 或 `{字段名}`

Q: PNG转换失败？
A: 确保系统已安装LibreOffice，且PPTX/DOCX文件格式正确

Q: 如何调整PNG分辨率？
A: 在 `convert_to_png()` 方法中修改 `dpi` 参数，默认为300dpi（高清打印质量）

## 后续改进方向

1. 增加GUI界面
2. 支持批量导入多个Excel文件
3. 支持自定义模板映射配置
4. 添加日志记录功能
5. 支持更多文件格式（XLS、ODT等）
6. 并行处理以提高效率
7. 错误重试机制

---

开发时间: 2026年5月12日
