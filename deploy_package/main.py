#!/usr/bin/env python3
"""
PPTX和DOCX批量邀请函生成系统 - 主运行脚本
直接运行此脚本来批量生成邀请函
"""

import os
import sys
import subprocess

# 获取脚本所在目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VENV_PYTHON = os.path.join(SCRIPT_DIR, '.venv', 'bin', 'python3')

def main():
    print("=" * 50)
    print("PPTX/DOCX 批量邀请函生成系统")
    print("=" * 50)
    print()
    
    # 检查虚拟环境
    if not os.path.exists(VENV_PYTHON):
        print("❌ 错误：虚拟环境不存在")
        print("请运行以下命令创建虚拟环境:")
        print(f"  cd {SCRIPT_DIR}")
        print(f"  python3 -m venv .venv")
        return False
    
    # 检查必需文件
    print("检查必需文件...")
    required_files = [
        'guest_list.xlsx',
        'invitation_template.pptx',
        'gov_invitation_template.docx'
    ]
    
    for filename in required_files:
        filepath = os.path.join(SCRIPT_DIR, filename)
        if not os.path.exists(filepath):
            print(f"❌ 缺少文件: {filename}")
            return False
        print(f"✓ 找到: {filename}")
    
    print()
    
    # 检查并安装依赖
    print("检查Python依赖...")
    pip_path = os.path.join(SCRIPT_DIR, '.venv', 'bin', 'pip')
    requirements_file = os.path.join(SCRIPT_DIR, 'requirements.txt')
    
    try:
        subprocess.run(
            [pip_path, 'install', '-q', '-r', requirements_file],
            check=False,
            capture_output=True,
            timeout=120
        )
        print("✓ 依赖检查完成")
    except Exception as e:
        print(f"⚠ 依赖安装可能不完整: {e}")
    
    print()
    print("=" * 50)
    print("开始处理邀请函...")
    print("=" * 50)
    print()
    
    # 设置Python路径
    env = os.environ.copy()
    env['PYTHONPATH'] = os.path.join(SCRIPT_DIR, 'src')
    
    # 运行主程序
    try:
        result = subprocess.run(
            [VENV_PYTHON, os.path.join(SCRIPT_DIR, 'src', 'batch_processor.py')],
            cwd=SCRIPT_DIR,
            env=env,
            timeout=300
        )
        
        if result.returncode == 0:
            print()
            print("=" * 50)
            print("✓ 处理成功完成！")
            print("=" * 50)
            return True
        else:
            print()
            print("❌ 处理失败，请查看上面的错误信息")
            return False
            
    except subprocess.TimeoutExpired:
        print("❌ 处理超时")
        return False
    except Exception as e:
        print(f"❌ 执行出错: {e}")
        return False


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
