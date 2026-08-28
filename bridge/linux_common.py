# -*- coding: utf-8 -*-
"""
linux_common.py — Linux 平台共享工具

职责：
1. 检测 WPS Office Linux CLI（et / wpp / wps）
2. 通过 CLI 打开文件（openWorkbook / openPresentation / openDocument）
3. 通过 CLI 转换 PDF（各 WPS 版本能力不一，失败时返回明确提示）
4. vendor 目录路径管理（openpyxl）

设计原则：纯 Python 标准库，x86/ARM 通用，无外网依赖。
"""
import os
import sys
import subprocess
from shutil import which

# vendor 目录（skill 根目录下 vendor/）
_VENDOR_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "vendor")


def vendor_path_ready() -> bool:
    """vendor 目录是否存在且包含 openpyxl"""
    return os.path.isdir(os.path.join(_VENDOR_DIR, "openpyxl"))


def setup_vendor_path():
    """把 vendor 目录加入 sys.path（幂等），供 openpyxl 导入"""
    if _VENDOR_DIR not in sys.path:
        sys.path.insert(0, _VENDOR_DIR)


def _find_wps_binary(name: str) -> str or None:
    """查找 WPS 二进制（et/wpp/wps），返回绝对路径或 None"""
    # 1. PATH 中查找
    p = which(name)
    if p:
        return p
    # 2. 常见安装路径
    candidates = [
        f"/usr/bin/{name}",
        f"/usr/local/bin/{name}",
        f"/opt/kingsoft/wps-office/office6/{name}",
        f"/opt/wps-office/office6/{name}",
        f"/snap/bin/wps-office-{name}",
    ]
    for c in candidates:
        if os.path.isfile(c) and os.access(c, os.X_OK):
            return c
    return None


class WpsCli:
    """WPS Office Linux CLI 封装"""

    def __init__(self):
        self.et_path = _find_wps_binary("et")      # 表格
        self.wpp_path = _find_wps_binary("wpp")    # 演示
        self.wps_path = _find_wps_binary("wps")    # 文字

    @property
    def available(self) -> bool:
        return any([self.et_path, self.wpp_path, self.wps_path])

    def info(self) -> dict:
        return {
            "et": self.et_path,       # WPS 表格
            "wpp": self.wpp_path,     # WPS 演示
            "wps": self.wps_path,     # WPS 文字
        }

    def open_file(self, app: str, file_path: str) -> dict:
        """用 Wps CLI 打开文件（GUI 弹出）"""
        if not file_path or not os.path.isfile(os.path.abspath(file_path)):
            return {"success": False, "error": f"文件不存在: {file_path}"}
        binary = {"excel": self.et_path, "ppt": self.wpp_path, "word": self.wps_path}.get(app)
        if not binary:
            return {"success": False, "error": f"未找到 WPS {app} CLI（et/wpp/wps），请确认 WPS Office for Linux 已安装"}
        try:
            subprocess.Popen(
                [binary, os.path.abspath(file_path)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return {"success": True, "data": {"file": os.path.abspath(file_path), "openedBy": binary}}
        except Exception as e:
            return {"success": False, "error": f"启动 WPS 失败: {e}"}

    def convert_to_pdf(self, app: str, file_path: str, out_path: str = None) -> dict:
        """尝试 CLI 转 PDF。注意：多数 WPS Linux 版本 CLI 不支持转换，需要 GUI 或 LibreOffice。"""
        if not file_path or not os.path.isfile(os.path.abspath(file_path)):
            return {"success": False, "error": f"文件不存在: {file_path}"}
        binary = {"excel": self.et_path, "ppt": self.wpp_path, "word": self.wps_path}.get(app)
        if not binary:
            return {"success": False, "error": "未找到 WPS CLI"}

        src = os.path.abspath(file_path)
        out = os.path.abspath(out_path) if out_path else os.path.splitext(src)[0] + ".pdf"

        # 尝试 WPS CLI 的 convert-to（部分版本支持）
        try:
            r = subprocess.run(
                [binary, "--convert-to", "pdf", "--outdir", os.path.dirname(out), src],
                capture_output=True, text=True, timeout=60,
            )
            if os.path.isfile(out):
                return {"success": True, "data": {"pdf": out}}
        except Exception:
            pass

        return {
            "success": False,
            "error": "WPS Linux CLI 不支持无界面转 PDF。请安装 LibreOffice（soffice --headless）或在 WPS GUI 中手动导出。",
        }


def linux_env_report() -> dict:
    """Linux 环境自检报告"""
    cli = WpsCli()
    return {
        "platform": "linux",
        "wps_cli": cli.info(),
        "wps_cli_available": cli.available,
        "vendor_ready": vendor_path_ready(),
        "vendor_dir": _VENDOR_DIR,
    }
