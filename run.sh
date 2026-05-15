#!/bin/bash

# PPTX和DOCX批量邀请函生成系统 - 运行脚本

set -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
VENV_PYTHON="$SCRIPT_DIR/.venv/bin/python3"

echo "=========================================="
echo "PPTX/DOCX 批量邀请函生成系统"
echo "=========================================="
echo ""

# 检查虚拟环境
if [ ! -f "$VENV_PYTHON" ]; then
    echo "❌ 错误：虚拟环境不存在，请先运行:"
    echo "  python3 -m venv .venv"
    exit 1
fi

# 检查依赖
echo "检查Python依赖..."
PIP="$SCRIPT_DIR/.venv/bin/pip"

# 安装requirements
echo "安装依赖包..."
$PIP install -q -r "$SCRIPT_DIR/requirements.txt" 2>/dev/null || {
    echo "⚠ 依赖安装可能不完整，继续..."
}

# 执行主程序
echo ""
echo "开始处理邀请函..."
echo "=========================================="
echo ""

cd "$SCRIPT_DIR"
export PYTHONPATH="$SCRIPT_DIR/src:$PYTHONPATH"

$VENV_PYTHON << 'PYTHON_SCRIPT'
import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from batch_processor import BatchProcessor

# 获取脚本目录
script_dir = os.path.dirname(os.path.abspath(__file__))

# 文件路径
excel_file = os.path.join(script_dir, 'guest_list.xlsx')
pptx_template = os.path.join(script_dir, 'invitation_template.pptx')
docx_template = os.path.join(script_dir, 'gov_invitation_template.docx')
output_docs = os.path.join(script_dir, 'output_documents')
output_images = os.path.join(script_dir, 'output_images')

# 验证文件
for file_path in [excel_file, pptx_template, docx_template]:
    if not os.path.exists(file_path):
        print(f"❌ 错误: 文件不存在 - {file_path}")
        sys.exit(1)

# 创建处理器并执行
processor = BatchProcessor(excel_file, pptx_template, docx_template, output_docs, output_images)

try:
    processor.process(generate_png=True)
except Exception as e:
    print(f"\n❌ 处理出错: {e}")
    sys.exit(1)

PYTHON_SCRIPT

echo ""
echo "=========================================="
echo "✓ 处理完成！"
echo "=========================================="
