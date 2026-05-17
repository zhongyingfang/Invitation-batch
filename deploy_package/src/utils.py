"""跨平台查找 LibreOffice / soffice 可执行文件"""
import os
import sys
import shutil
import subprocess


def find_soffice() -> str | None:
    """返回 soffice 可执行文件的完整路径，找不到返回 None。"""
    # 1) 环境变量 LO_PATH / LIBREOFFICE_PATH 优先
    for env_name in ("LO_PATH", "LIBREOFFICE_PATH"):
        val = os.environ.get(env_name)
        if val and os.path.isfile(val):
            return val
        if val and os.path.isdir(val):
            for name in ("soffice.exe", "soffice"):
                p = os.path.join(val, name)
                if os.path.isfile(p):
                    return p

    # 2) PATH 里找
    try:
        for name in ("soffice", "libreoffice"):
            found = shutil.which(name)
            if found:
                return found
    except Exception:
        pass

    # 3) Windows 常见安装路径 + 注册表
    if sys.platform == "win32":
        # 常见路径
        for base in (
            os.environ.get("ProgramFiles", "C:\\Program Files"),
            os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"),
        ):
            p = os.path.join(base, "LibreOffice", "program", "soffice.exe")
            if os.path.isfile(p):
                return p

        # 注册表
        import winreg  # 仅 Windows 可用
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                key = winreg.OpenKey(root, r"Software\LibreOffice\UNO\InstallPath")
                val, _ = winreg.QueryValueEx(key, "")
                winreg.CloseKey(key)
                p = os.path.join(val, "soffice.exe")
                if os.path.isfile(p):
                    return p
            except Exception:
                continue

    # 4) macOS 常见路径
    if sys.platform == "darwin":
        mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.isfile(mac_path):
            return mac_path

    # 5) Linux 常见路径 (Docker 容器)
    for linux_path in (
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice",
        "/opt/libreoffice*/program/soffice",
    ):
        if os.path.isfile(linux_path):
            return linux_path
        import glob
        matches = glob.glob(linux_path)
        for m in matches:
            if os.path.isfile(m):
                return m

    return None
