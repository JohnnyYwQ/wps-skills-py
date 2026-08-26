#!/usr/bin/env python3
"""
WPS 统一 Skill 安装脚本（Excel + PPT + Word）
检查环境依赖并配置服务。

用法：
  python install.py              # 安装
  python install.py --check      # 仅检查环境
"""

import sys
import os
import platform
import json

def check_python():
    """检查 Python 版本"""
    v = sys.version_info
    print(f"  Python: {v.major}.{v.minor}.{v.micro} ({platform.architecture()[0]})")
    if v.major < 3 or (v.major == 3 and v.minor < 8):
        print("  [警告] 需要 Python 3.8+，当前版本可能不兼容")
        return False
    print("  [OK] Python 版本符合要求")
    return True


def check_platform():
    """检查平台"""
    s = platform.system()
    a = platform.machine()
    print(f"  平台: {s} {a}")
    if s == "Windows":
        print("  [OK] Windows 平台（通过 PowerShell COM 控制 WPS，233 个 action）")
        return True
    elif s == "Linux":
        if a in ("x86_64", "aarch64", "arm64"):
            print(f"  [OK] Linux {a} 平台（文件级 OpenXML 后端：Excel 56 / PPT 20 / Word 30 个 action）")
            return True
        print(f"  [提示] Linux {a} 架构未在验证列表（x86_64/aarch64）中，理论上纯 Python 后端仍可用")
        return True
    else:
        print(f"  [警告] 平台 {s} 不支持（macOS 无 WPS 自动化接口）")
        return False


def _check_progid(progid):
    """检查指定 WPS 应用的 COM ProgID 是否在注册表注册"""
    try:
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rf"Software\Classes\{progid}\CLSID") as key:
                return winreg.QueryValue(key, "")
        except FileNotFoundError:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Classes\{progid}\CLSID") as key:
                return winreg.QueryValue(key, "")
    except Exception:
        return None

def check_wps():
    """检查 WPS Office 三应用 COM 是否注册"""
    s = platform.system()
    if s == "Windows":
        progs = {
            "Ket.Application": "WPS 表格(Excel)",
            "Kwpp.Application": "WPS 演示(PPT)",
            "Kwps.Application": "WPS 文字(Word)",
        }
        ok = True
        for progid, label in progs.items():
            clsid = _check_progid(progid)
            if clsid:
                print(f"  [OK] {label} COM: {progid} (CLSID: {clsid})")
            else:
                print(f"  [警告] 未找到 {label} COM 注册({progid})，请确认 WPS 已安装")
                ok = False
        return ok
    elif s == "Linux":
        import shutil
        found = []
        for name in ("et", "wpp", "wps"):
            p = shutil.which(name)
            if p:
                found.append(f"{name}={p}")
        # WPS GUI 在 Linux 上仅用于预览（openInWps），非必需
        if found:
            print(f"  [OK] WPS 命令行可用: {', '.join(found)}（用于 GUI 打开预览）")
        else:
            # 检查常见安装路径
            alt = [p for p in ("/usr/bin/et", "/opt/kingsoft/wps-office/office6/et") if os.path.exists(p)]
            if alt:
                print(f"  [OK] 找到 WPS: {alt[0]}（用于 GUI 打开预览）")
            else:
                print("  [提示] 未找到 WPS 命令行（GUI 预览不可用；文件级自动化不受影响，无需 WPS）")
        return True
    return False


def check_files():
    """检查必要文件"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)

    files = [
        ("SKILL.md", "技能定义文件"),
        ("bridge/action_trace.py", "Action 结构化追踪与 24 小时保留"),
        ("bridge/server.py", "统一桥接服务"),
        ("bridge/wps_excel.py", "WPS Excel 控制器"),
        ("bridge/wps_ppt.py", "WPS PPT 控制器"),
        ("bridge/wps_word.py", "WPS Word 控制器"),
        ("bridge/linux_common.py", "Linux 公共后端（vendor 路径 + WPS CLI）"),
        ("bridge/linux_excel.py", "Linux Excel 后端（openpyxl）"),
        ("bridge/linux_ppt.py", "Linux PPT 后端（OpenXML）"),
        ("bridge/linux_word.py", "Linux Word 后端（OpenXML）"),
        ("vendor/openpyxl/__init__.py", "vendored openpyxl（Linux Excel 后端依赖，无需 pip）"),
        ("vendor/et_xmlfile/__init__.py", "vendored et_xmlfile（openpyxl 依赖）"),
        ("config.json", "配置文件"),
    ]

    all_ok = True
    for rel, desc in files:
        path = os.path.join(project_dir, rel)
        if os.path.exists(path):
            size = os.path.getsize(path)
            print(f"  [OK] {rel} ({size} bytes) - {desc}")
        else:
            print(f"  [缺失] {rel} - {desc}")
            all_ok = False
    return all_ok


def show_usage():
    """显示使用说明"""
    print()
    print("=" * 60)
    print("  使用说明")
    print("=" * 60)
    print()
    print("  1. 调用 Action（会自动启动桥接服务）：")
    print("     python scripts/call.py getContext '{}'")
    print()
    print("  2. 如需前台运行桥接服务：")
    print("     python scripts/start.py")
    print()
    print("  3. 测试连接：")
    print("     python scripts/test.py")
    print()
    print("  4. HTTP POST 调用方式：")
    print("     POST http://127.0.0.1:58891/dispatch")
    print('     {"action": "getContext", "params": {}}')
    print()
    print()
    print("=" * 60)


def main():
    print()
    print("=" * 60)
    print("  WPS 统一 Skill 环境检查（Excel + PPT + Word）")
    print("  纯 Python | 无 MCP | 无外网 | 无 Node.js")
    print("=" * 60)
    print()

    print("[1/4] 检查 Python 环境")
    check_python()
    print()

    print("[2/4] 检查平台")
    check_platform()
    print()

    print("[3/4] 检查 WPS Office")
    check_wps()
    print()

    print("[4/4] 检查文件完整性")
    check_files()
    print()

    show_usage()


if __name__ == "__main__":
    main()
