# -*- coding: utf-8 -*-
"""
test_linux_sim.py — 在 Windows 上模拟 Linux 分支，验证控制器 -> Linux 桥接的全链路委托
方法：临时把 wps_excel / wps_ppt / wps_word 模块的 IS_LINUX/IS_WINDOWS 反转后实例化控制器。
用途：无 Linux 真机时的 Linux 路径回归测试（26 项断言）。
"""
import os
import sys
import json

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BRIDGE = os.path.join(ROOT, "bridge")
sys.path.insert(0, BRIDGE)

OUT = os.path.join(os.environ.get("TEMP", ROOT))
XLSX = os.path.join(OUT, "_sim_linux.xlsx")
PPTX = os.path.join(OUT, "_sim_linux.pptx")
DOCX = os.path.join(OUT, "_sim_linux.docx")

passed, failed = [], []

def check(name, cond, detail=""):
    (passed if cond else failed).append(name)
    print(("PASS" if cond else "FAIL") + f"  {name}" + (f"  [{detail}]" if detail else ""))

def make_ctrl(mod_name):
    import importlib
    mod = importlib.import_module(mod_name)
    mod.IS_LINUX = True
    mod.IS_WINDOWS = False
    return mod, mod.WpsExcelController() if hasattr(mod, "WpsExcelController") else \
        (mod.WpsPptController() if hasattr(mod, "WpsPptController") else mod.WpsWordController())

def run(ctrl, action, **params):
    return ctrl.execute(action, params)

# ---------- Excel ----------
print("\n=== Excel 控制器 -> LinuxExcelBridge ===")
mod, ctrl = make_ctrl("wps_excel")
r = run(ctrl, "ping")
check("excel ping", r.get("success") and "openpyxl" in str(r.get("data", {})), str(r))
r = run(ctrl, "createWorkbook", filePath=XLSX)
check("excel createWorkbook", r.get("success"), str(r))
r = run(ctrl, "setCell", row=1, col=1, value="名称") if run(ctrl, "ping") else None
# 用实际 action 名（与 Windows 契约一致）
r = run(ctrl, "setCellText", row=1, col=1, text="名称")
if not r.get("success"):
    r = run(ctrl, "setCellValue", row=1, col=1, value="名称")
check("excel setCellValue", r.get("success"), str(r))
r = run(ctrl, "setCellFormula", cell="B2", formula="=1+1") if run(ctrl, "ping") else None
r = run(ctrl, "setFormula", cell="B2", formula="=1+1")
check("excel setFormula", r.get("success"), str(r))
r = run(ctrl, "save")
check("excel save", r.get("success") and os.path.isfile(XLSX), str(r))
r = run(ctrl, "openWorkbook", filePath=XLSX)
check("excel reopen", r.get("success"), str(r))
r = run(ctrl, "getCellValue", cell="B2")
check("excel getCellValue formula roundtrip", r.get("success"), str(r))
r = run(ctrl, "getSelectedText")
check("excel unsupported action honest error", not r.get("success") and "Linux" in str(r.get("error", "")), str(r))

# ---------- PPT ----------
print("\n=== PPT 控制器 -> LinuxPptBridge ===")
mod, ctrl = make_ctrl("wps_ppt")
r = run(ctrl, "ping")
check("ppt ping", r.get("success"), str(r))
r = run(ctrl, "createPresentation", filePath=PPTX)
check("ppt createPresentation", r.get("success"), str(r))
r = run(ctrl, "addSlide", layout="title_content")
check("ppt addSlide", r.get("success"), str(r))
r = run(ctrl, "setSlideTitle", index=1, title="模拟测试")
check("ppt setSlideTitle", r.get("success"), str(r))
r = run(ctrl, "setSlideNotes", slideIndex=1, text="备注内容")
check("ppt setSlideNotes", r.get("success"), str(r))
r = run(ctrl, "getSlideCount")
check("ppt getSlideCount", r.get("success") and r.get("data", {}).get("slideCount") == 1, str(r))
r = run(ctrl, "save")
check("ppt save", r.get("success") and os.path.isfile(PPTX), str(r))
r = run(ctrl, "openPresentation", filePath=PPTX)
check("ppt reopen", r.get("success"), str(r))
r = run(ctrl, "getSlideNotes", slideIndex=1)
check("ppt notes roundtrip", r.get("success") and "备注" in str(r.get("data", {})), str(r))

# ---------- Word ----------
print("\n=== Word 控制器 -> LinuxWordBridge ===")
mod, ctrl = make_ctrl("wps_word")
r = run(ctrl, "ping")
check("word ping", r.get("success"), str(r))
r = run(ctrl, "createDocument", filePath=DOCX)
check("word createDocument", r.get("success"), str(r))
r = run(ctrl, "insertText", text="跨平台标题", style="Title")
check("word insertText Title", r.get("success"), str(r))
r = run(ctrl, "insertText", text="正文段落内容", style="Normal")
check("word insertText Normal", r.get("success"), str(r))
r = run(ctrl, "insertTable", rows=2, cols=2, data=[["a", "b"], ["c", "d"]])
check("word insertTable", r.get("success"), str(r))
r = run(ctrl, "getDocumentParagraphs")
check("word getDocumentParagraphs", r.get("success"), str(r))
r = run(ctrl, "save")
check("word save", r.get("success") and os.path.isfile(DOCX), str(r))
r = run(ctrl, "openDocument", filePath=DOCX)
check("word reopen", r.get("success"), str(r))
r = run(ctrl, "getDocumentText")
check("word text roundtrip", r.get("success") and "跨平台标题" in str(r.get("data", "")), str(r))

print(f"\n总计: {len(passed)} 通过, {len(failed)} 失败")
if failed:
    print("失败项:", failed)
sys.exit(1 if failed else 0)
