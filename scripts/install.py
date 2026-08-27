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

BRIDGE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "bridge")
if BRIDGE_DIR not in sys.path:
    sys.path.insert(0, BRIDGE_DIR)

from windows_com import resolve_com_runtime

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
        print("  [OK] Windows 平台（PowerShell COM：233 个应用 Action + 2 个 bridge Contract）")
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
            resolution = resolve_com_runtime(progid)
            if resolution.available:
                registration = resolution.selected_registration
                print(
                    f"  [OK] {label} COM: {progid} "
                    f"({resolution.selected_view_bitness}-bit view; "
                    f"CLSID: {registration.clsid}; "
                    f"{registration.activation_kind}; "
                    f"PowerShell: {resolution.powershell_executable})"
                )
            else:
                print(
                    f"  [警告] {label} COM 注册不可激活({progid}): "
                    f"{resolution.diagnostic}"
                )
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
        ("bridge/action_catalog.py", "Action Contract Catalog 与校验器"),
        ("bridge/action_manifest.json", "Windows Action Contract 唯一事实源"),
        ("bridge/powershell_contracts.py", "共享 PowerShell 契约枚举转换"),
        ("bridge/service_lifecycle.py", "bridge 实例身份与生命周期"),
        ("bridge/windows_com.py", "Windows COM 注册视图与 PowerShell 解析"),
        ("bridge/server.py", "统一桥接服务"),
        ("bridge/wps_excel.py", "WPS Excel 控制器"),
        ("bridge/wps_ppt.py", "WPS PPT 控制器"),
        ("bridge/wps_word.py", "WPS Word 控制器"),
        ("bridge/linux_common.py", "Linux 公共后端（vendor 路径 + WPS CLI）"),
        ("bridge/linux_excel.py", "Linux Excel 后端（openpyxl）"),
        ("bridge/linux_ppt.py", "Linux PPT 后端（OpenXML）"),
        ("bridge/linux_word.py", "Linux Word 后端（OpenXML）"),
        ("scripts/call.py", "Action 调用入口"),
        ("scripts/actions.py", "Action Contract 查询入口"),
        ("scripts/service.py", "bridge 生命周期命令"),
        ("vendor/openpyxl/__init__.py", "vendored openpyxl（Linux Excel 后端依赖，无需 pip）"),
        ("vendor/et_xmlfile/__init__.py", "vendored et_xmlfile（openpyxl 依赖）"),
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
    print("  2. 查看或停止桥接服务：")
    print("     python scripts/service.py status")
    print("     python scripts/service.py stop")
    print()
    print("  3. 测试连接：")
    print("     python scripts/test.py")
    print()
    print("  4. 如需前台观察服务：")
    print("     python scripts/start.py")
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
