#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""统一 WPS Skill 功能验证脚本（Excel / PPT / Word + 通用）。
通过 scripts/call.py 的胶水层调用，会自动拉起最新桥接服务。"""
import sys, os, json, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import call  # noqa: E402

PORT = call.PORT
_results = []

def run(label, action, params=None):
    params = params or {}
    if not call._ensure_server():
        _results.append((label, False, {"error": "server start failed"}))
        return
    try:
        r = call._post(action, params)
    except Exception as e:
        r = {"success": False, "error": str(e)}
    ok = bool(r.get("success"))
    _results.append((label, ok, r))
    tag = "OK  " if ok else "FAIL"
    print(f"[{tag}] {label}: {json.dumps(r, ensure_ascii=False)}")
    time.sleep(0.3)
    return r

print("=== 健康检查 ===")
h = call._health()
print(json.dumps(h, ensure_ascii=False))

print("\n=== Excel 功能 ===")
run("Excel createWorkbook", "createWorkbook", {})
run("Excel setCellValue(1,1)", "setCellValue", {"row": 1, "col": 1, "value": "hello统一"})
run("Excel getCellValue(1,1)", "getCellValue", {"row": 1, "col": 1})
run("Excel setFormula(D2)", "setFormula", {"cell": "D2", "formula": "=1+2"})
r_getf = run("Excel getFormula(D2)", "getFormula", {"cell": "D2"})
run("Excel getActiveWorkbook", "getActiveWorkbook", {})

print("\n=== PPT 功能 ===")
run("PPT createPresentation", "createPresentation", {})
run("PPT addSlide(title)", "addSlide", {"layout": "title_content", "title": "测试标题"})
run("PPT getSlideCount", "getSlideCount", {})
run("PPT getSlideTitle(1)", "getSlideTitle", {"slideIndex": 1})
run("PPT setSlideTitle(1)", "setSlideTitle", {"slideIndex": 1, "title": "新标题统一"})
run("PPT getSlideTitle(1) after set", "getSlideTitle", {"slideIndex": 1})

print("\n=== Word 功能 ===")
run("Word createDocument", "createDocument", {})
run("Word insertText", "insertText", {"text": "这是一段测试文字统一", "position": "end"})
run("Word getDocumentText", "getDocumentText", {})
run("Word getActiveDocument", "getActiveDocument", {})
run("Word setFont", "setFont", {"font_name": "微软雅黑", "font_size": 14})

print("\n=== 通用 action（按 app 委派）===")
run("Common getAppInfo(excel)", "getAppInfo", {"app": "excel"})
run("Common getAppInfo(ppt)", "getAppInfo", {"app": "ppt"})
run("Common getAppInfo(word)", "getAppInfo", {"app": "word"})

print("\n=== 汇总 ===")
ok_n = sum(1 for _, ok, _ in _results if ok)
fail_n = len(_results) - ok_n
print(f"通过 {ok_n}/{len(_results)}，失败 {fail_n}")
for label, ok, r in _results:
    if not ok:
        print(f"  - FAIL: {label} -> {r.get('error', r)}")
sys.exit(0 if fail_n == 0 else 1)
