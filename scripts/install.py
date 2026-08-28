#!/usr/bin/env python3
"""Read-only environment report for the WPS Skill."""

import argparse
import os
import platform
import shutil
import sys


BRIDGE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "bridge")
if BRIDGE_DIR not in sys.path:
    sys.path.insert(0, BRIDGE_DIR)

from windows_com import resolve_com_runtime


def check_python():
    version = sys.version_info
    print(f"  Python: {version.major}.{version.minor}.{version.micro} ({platform.architecture()[0]})")
    if version < (3, 8):
        print("  [警告] Action Runtime 需要 Python 3.8 或更高版本")
        return False
    print("  [OK] Python 版本符合 Runtime 要求")
    return True


def check_platform():
    system = platform.system()
    print(f"  平台: {system} {platform.machine()}")
    if system == "Windows":
        print("  [OK] Windows 实机运行环境")
    elif system == "Linux":
        print("  [提示] Linux 文件后端保留在仓库中，但不属于本轮实机验收范围")
    else:
        print("  [提示] 当前平台不能执行 Windows WPS COM Action；仍可运行非 WPS 自动化套件")
    return True


def check_windows_powershell():
    if platform.system() != "Windows":
        print("  [提示] 非 Windows：跳过 Windows PowerShell 检查")
        return True

    executable = shutil.which("powershell.exe") or shutil.which("powershell")
    if executable:
        print(f"  [OK] Windows PowerShell: {executable}")
        return True
    print("  [缺失] Windows PowerShell：未在 PATH 中找到 powershell.exe")
    return False


def check_windows_powershell_and_wps():
    if platform.system() != "Windows":
        print("  [提示] 非 Windows：跳过 Windows PowerShell 与 WPS COM 注册检查")
        return True

    progids = {
        "Ket.Application": "WPS 表格 (Excel)",
        "Kwpp.Application": "WPS 演示 (PPT)",
        "Kwps.Application": "WPS 文字 (Word)",
    }
    all_available = True
    for progid, label in progids.items():
        resolution = resolve_com_runtime(progid)
        if resolution.available:
            registration = resolution.selected_registration
            print(
                f"  [OK] {label}: {progid} ({resolution.selected_view_bitness}-bit; "
                f"CLSID: {registration.clsid}; PowerShell: {resolution.powershell_executable}; "
                f"{registration.activation_kind})"
            )
        else:
            print(f"  [缺失] {label}: {progid} ({resolution.diagnostic})")
            all_available = False
    return all_available


def check_resources():
    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    resources = (
        ("SKILL.md", "Skill 操作说明"),
        ("bridge/action_manifest.json", "Action Contract 唯一事实源"),
        ("bridge/action_catalog.py", "Catalog"),
        ("bridge/action_runtime.py", "Action Runtime"),
        ("bridge/action_gate.py", "Windows Action mutex"),
        ("bridge/action_trace.py", "Action trace"),
        ("scripts/actions.py", "Catalog CLI"),
        ("scripts/call.py", "Action CLI"),
        ("vendor/openpyxl/__init__.py", "Linux Excel vendored dependency"),
        ("vendor/et_xmlfile/__init__.py", "vendored dependency"),
    )
    all_present = True
    for relative_path, description in resources:
        if os.path.isfile(os.path.join(project_dir, relative_path)):
            print(f"  [OK] {relative_path} - {description}")
        else:
            print(f"  [缺失] {relative_path} - {description}")
            all_present = False
    return all_present


def main(argv=None):
    parser = argparse.ArgumentParser(description="报告 WPS Skill 本地环境，不修改系统")
    parser.add_argument("--check", action="store_true", help="保留的显式检查标志")
    parser.parse_args(argv)

    print("WPS Skill 环境报告（只读）")
    print("本检查不会修改系统、下载内容或创建运行环境。")
    print()
    print("[1/5] Python")
    check_python()
    print("[2/5] 平台")
    check_platform()
    print("[3/5] Windows PowerShell")
    check_windows_powershell()
    print("[4/5] WPS COM 注册")
    check_windows_powershell_and_wps()
    print("[5/5] 仓内资源")
    check_resources()
    print()
    print("下一步：先运行 scripts/actions.py 查询 Contract，再用 scripts/call.py 执行一个 Action。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
