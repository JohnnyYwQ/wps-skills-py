#!/usr/bin/env python3
"""Representative real-WPS checks through the public one-Action CLI."""

import json
from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parents[1]
CALL = ROOT / "scripts" / "call.py"


def run_action(label, action, params=None, app=None):
    command = [sys.executable, str(CALL), action]
    if app:
        command.extend(("--app", app))
    command.append(json.dumps(params or {}, ensure_ascii=False))
    completed = subprocess.run(
        command,
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    try:
        response = json.loads(completed.stdout)
    except json.JSONDecodeError:
        response = {
            "success": False,
            "error": f"Action CLI did not return JSON: {completed.stderr or completed.stdout}",
        }
    ok = bool(response.get("success"))
    print(f"[{'OK  ' if ok else 'FAIL'}] {label}: {json.dumps(response, ensure_ascii=False)}")
    return label, ok, response


def main():
    results = []

    print("=== Excel 功能 ===")
    results.append(run_action("Excel createWorkbook", "createWorkbook", app="excel"))
    results.append(run_action(
        "Excel setCellValue(1,1)", "setCellValue",
        {"row": 1, "col": 1, "value": "hello统一"}, "excel",
    ))
    results.append(run_action(
        "Excel getCellValue(1,1)", "getCellValue", {"row": 1, "col": 1}, "excel",
    ))
    results.append(run_action(
        "Excel setFormula(D2)", "setFormula", {"cell": "D2", "formula": "=1+2"}, "excel",
    ))

    print("\n=== PPT 功能 ===")
    results.append(run_action("PPT createPresentation", "createPresentation", app="ppt"))
    results.append(run_action(
        "PPT addSlide(title)", "addSlide",
        {"layout": "title_content", "title": "测试标题"}, "ppt",
    ))
    results.append(run_action("PPT getSlideCount", "getSlideCount", app="ppt"))

    print("\n=== Word 功能 ===")
    results.append(run_action("Word createDocument", "createDocument", app="word"))
    results.append(run_action(
        "Word insertText", "insertText", {"text": "这是一段测试文字统一", "position": "end"}, "word",
    ))
    results.append(run_action("Word getDocumentText", "getDocumentText", app="word"))

    print("\n=== 应用连接 ===")
    for app in ("excel", "ppt", "word"):
        results.append(run_action(f"{app} getAppInfo", "getAppInfo", app=app))

    passed = sum(1 for _, ok, _ in results if ok)
    failed = len(results) - passed
    print(f"\n通过 {passed}/{len(results)}，失败 {failed}")
    for label, ok, response in results:
        if not ok:
            print(f"  - FAIL: {label} -> {response.get('error', response)}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
