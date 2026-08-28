#!/usr/bin/env python3
"""Run all Action Contracts against real WPS Office on Windows.

This is a WPS task orchestrator for validation only.  Each Action still travels
through the public ``scripts/call.py`` one-Action CLI and therefore exercises the
same Action Runtime, execution bridge, PowerShell bridge, and WPS COM behavior as
a real caller.
"""

import argparse
import base64
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time
import zipfile


ROOT = Path(__file__).resolve().parents[1]
BRIDGE_DIR = ROOT / "bridge"
if str(BRIDGE_DIR) not in sys.path:
    sys.path.insert(0, str(BRIDGE_DIR))

from action_catalog import ActionCatalog  # noqa: E402
from windows_wps_full_suite import (  # noqa: E402
    SCENARIOS,
    build_action_params,
    default_suite_state,
    iter_declared_actions,
)


CALL_SCRIPT = ROOT / "scripts" / "call.py"
INTERACTIVE_ACTIONS = {
    ("ppt", "getSelectedText"),
    ("ppt", "setSelectedText"),
    ("ppt", "startSlideShow"),
}
FILE_EFFECTS = {
    ("excel", "saveAs"): "filePath",
    ("excel", "convertToPDF"): "outputPath",
    ("excel", "convertFormat"): "outputPath",
    ("excel", "exportChartAsImage"): "outputPath",
    ("excel", "exportRangeAsImage"): "outputPath",
    ("ppt", "saveAs"): "filePath",
    ("ppt", "convertToPDF"): "outputPath",
    ("ppt", "convertFormat"): "outputPath",
    ("ppt", "exportSlideAsImage"): "outputPath",
    ("word", "saveAs"): "filePath",
    ("word", "convertToPDF"): "outputPath",
    ("word", "convertFormat"): "outputPath",
}


def _utc_now():
    return datetime.now(timezone.utc).isoformat()


def _write_fixture_png(path: Path):
    # Valid 1x1 RGBA PNG; no third-party image library is needed on the test PC.
    encoded = (
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4"
        "z8DwHwAFgAI/ScL5WQAAAABJRU5ErkJggg=="
    )
    path.write_bytes(base64.b64decode(encoded))


def _write_minimal_docx(path: Path, text: str):
    """Create a small standards-based DOCX fixture using only the stdlib."""

    escaped = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
    relationships = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>{escaped}</w:t></w:r></w:p><w:sectPr/></w:body>
</w:document>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", relationships)
        archive.writestr("word/document.xml", document)


def validate_suite(catalog: ActionCatalog, state: dict):
    declared = list(iter_declared_actions())
    manifest = {
        (item["owner"], item["action"])
        for item in catalog.list()
    }
    if len(declared) != len(set(declared)):
        raise ValueError("full suite contains duplicate Action identities")
    if set(declared) != manifest:
        missing = sorted(manifest - set(declared))
        unknown = sorted(set(declared) - manifest)
        raise ValueError(f"full suite coverage mismatch: missing={missing}; unknown={unknown}")
    for owner, action in declared:
        catalog.validate_params(
            owner,
            action,
            build_action_params(owner, action, state),
        )
    return len(declared)


def _response_data(response):
    if not isinstance(response, dict):
        return {}
    data = response.get("data")
    return data if isinstance(data, dict) else {}


def _effect_errors(owner: str, action: str, params: dict, response: dict):
    """Return deterministic effect/readback failures beyond Runtime validation."""

    errors = []
    data = _response_data(response)
    file_parameter = FILE_EFFECTS.get((owner, action))
    if file_parameter:
        output = Path(params[file_parameter])
        if not output.is_file():
            errors.append(f"output file was not created: {output}")
        elif output.stat().st_size <= 0:
            errors.append(f"output file is empty: {output}")

    if owner == "bridge":
        unavailable = [app for app in ("excel", "ppt", "word") if data.get(app) is not True]
        if unavailable:
            errors.append(f"WPS COM unavailable: {', '.join(unavailable)}")
    elif owner == "excel":
        if action == "getCellValue" and data.get("value") != 42:
            errors.append(f"expected cell value 42, got {data.get('value')!r}")
        elif action == "getRangeData" and (data.get("rows"), data.get("cols")) != (5, 2):
            errors.append(f"expected a 5x2 range, got {data.get('rows')}x{data.get('cols')}")
        elif action == "getFormula" and "B2" not in str(data.get("formula", "")).upper():
            errors.append(f"formula readback does not reference B2: {data.get('formula')!r}")
        elif action == "getSelectedText" and "EXCEL_SELECTION_TOKEN" not in str(data.get("text", "")):
            errors.append(f"selected text readback mismatch: {data.get('text')!r}")
        elif action == "getCellComments" and len(data.get("comments", [])) < 1:
            errors.append("comment readback returned no comments")
        elif action == "findReplace" and int(data.get("count", 0)) < 1:
            errors.append("find/replace returned no replacements")
    elif owner == "ppt":
        if action == "getSlideTitle" and "PPT_TITLE_TOKEN" not in str(data.get("title", "")):
            errors.append(f"slide title readback mismatch: {data.get('title')!r}")
        elif action == "getSlideNotes" and "PPT_NOTES_TOKEN" not in str(data.get("notes", "")):
            errors.append(f"slide notes readback mismatch: {data.get('notes')!r}")
        elif action == "getPptTableCell" and "PPT_TABLE_TOKEN" not in str(data.get("value", "")):
            errors.append(f"table cell readback mismatch: {data.get('value')!r}")
        elif action == "getSelectedText" and "PPT_SELECTION_TOKEN" not in str(data.get("text", "")):
            errors.append(f"selected text readback mismatch: {data.get('text')!r}")
        elif action == "findPptText" and int(data.get("count", 0)) < 1:
            errors.append("text search returned no matches")
    elif owner == "word":
        if action == "getDocumentText" and "WORD_FIND_TOKEN" not in str(data.get("text", "")):
            errors.append("document text readback does not contain WORD_FIND_TOKEN")
        elif action == "getDocumentParagraphs" and int(data.get("count", 0)) < 1:
            errors.append("paragraph readback returned no paragraphs")
        elif action == "getSelectedText" and "WORD_SELECTION_TOKEN" not in str(data.get("text", "")):
            errors.append(f"selected text readback mismatch: {data.get('text')!r}")
        elif action == "findInDocument" and int(data.get("count", 0)) < 1:
            errors.append("document search returned no matches")
    return errors


class ActionRecorder:
    def __init__(self, output_dir: Path):
        self.output_dir = output_dir
        self.records = []
        self.jsonl_path = output_dir / "actions.jsonl"
        self.jsonl_path.write_text("", encoding="utf-8")
        self.trace_dir = output_dir / "traces"
        self.trace_dir.mkdir(parents=True, exist_ok=True)

    def append(self, record: dict):
        self.records.append(record)
        with self.jsonl_path.open("a", encoding="utf-8") as stream:
            stream.write(json.dumps(record, ensure_ascii=False, allow_nan=False) + "\n")

    def copy_trace(self, response: dict):
        trace_path = response.get("traceLog") if isinstance(response, dict) else None
        trace_id = response.get("traceId") if isinstance(response, dict) else None
        if not trace_path or not trace_id:
            return None
        source = Path(trace_path)
        if not source.is_file():
            return None
        target = self.trace_dir / f"{trace_id}.jsonl"
        try:
            shutil.copy2(source, target)
            return str(target)
        except OSError:
            return None


class LiveSuite:
    def __init__(self, output_dir: Path, *, skip_interactive: bool):
        self.output_dir = output_dir
        self.skip_interactive = skip_interactive
        self.state = default_suite_state(output_dir)
        self.catalog = ActionCatalog.from_path()
        self.recorder = ActionRecorder(output_dir)
        self.started_at = _utc_now()

    def _record_harness(self, label: str, success: bool, detail: str, duration_ms: float):
        status = "pass" if success else "fail"
        self.recorder.append({
            "scenario": "Harness precondition",
            "owner": "harness",
            "action": label,
            "role": "setup",
            "status": status,
            "durationMs": round(duration_ms, 2),
            "detail": detail,
        })
        print(f"[{'PASS' if success else 'FAIL'}] setup.{label}")
        return success

    def _powershell_setup(self, label: str, script: str):
        started = time.perf_counter()
        try:
            completed = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
                cwd=ROOT,
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                timeout=60,
                check=False,
            )
            output = completed.stdout.strip()
            detail = output or f"exit code {completed.returncode}"
            success = completed.returncode == 0
        except Exception as exc:
            success = False
            detail = f"{type(exc).__name__}: {exc}"
        return self._record_harness(
            label,
            success,
            detail,
            (time.perf_counter() - started) * 1000,
        )

    def _invoke(self, scenario: str, owner: str, action: str, params: dict, role: str):
        command = [sys.executable, "-X", "utf8", str(CALL_SCRIPT), action]
        if owner != "bridge":
            command.extend(["--app", owner])
        command.append("--stdin")
        child_env = os.environ.copy()
        child_env["PYTHONUTF8"] = "1"
        started = time.perf_counter()
        response = None
        stdout = ""
        exit_code = None
        transport_error = None
        try:
            completed = subprocess.run(
                command,
                cwd=ROOT,
                env=child_env,
                input=json.dumps(params, ensure_ascii=False),
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                timeout=750,
                check=False,
            )
            exit_code = completed.returncode
            stdout = completed.stdout.strip()
            try:
                response = json.loads(stdout)
            except json.JSONDecodeError as exc:
                transport_error = f"Action CLI returned invalid JSON: {exc}: {stdout}"
        except Exception as exc:
            transport_error = f"{type(exc).__name__}: {exc}"

        effect_errors = []
        response_success = bool(isinstance(response, dict) and response.get("success"))
        if response_success and role == "primary":
            effect_errors = _effect_errors(owner, action, params, response)
        passed = response_success and not effect_errors and transport_error is None
        copied_trace = self.recorder.copy_trace(response or {})
        record = {
            "scenario": scenario,
            "owner": owner,
            "action": action,
            "role": role,
            "status": "pass" if passed else "fail",
            "durationMs": round((time.perf_counter() - started) * 1000, 2),
            "params": params,
            "exitCode": exit_code,
            "response": response,
            "transportError": transport_error,
            "effectErrors": effect_errors,
            "copiedTrace": copied_trace,
        }
        self.recorder.append(record)
        print(f"[{'PASS' if passed else 'FAIL'}] {owner}.{action} ({role})")
        return response or {}, passed

    def _setup_action(self, owner: str, action: str, params: dict):
        return self._invoke("Harness fixture setup", owner, action, params, "setup")

    def _select_for_selected_text(self, owner: str):
        if owner == "excel":
            return self._powershell_setup(
                "select-excel-cell",
                "$app=[System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application'); "
                "$app.ActiveWorkbook.ActiveSheet.Range('A30').Select() | Out-Null",
            )
        if owner == "ppt":
            return self._powershell_setup(
                "select-ppt-title-text",
                "$app=[System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwpp.Application'); "
                "$shape=$app.ActivePresentation.Slides.Item(1).Shapes.Item(1); "
                "$shape.TextFrame.TextRange.Select()",
            )
        if owner == "word":
            return self._powershell_setup(
                "select-word-insertion-point",
                "$app=[System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application'); "
                "$doc=$app.ActiveDocument; $pos=[Math]::Max(0,$doc.Content.End-1); "
                "$doc.Range($pos,$pos).Select()",
            )
        return True

    def _before_primary(self, owner: str, action: str):
        if action == "setSelectedText":
            self._select_for_selected_text(owner)
        elif action == "getSelectedText" and owner == "ppt":
            # setSelectedText normally leaves the inserted text selected.  Re-select
            # the title if WPS collapsed its Selection object between Action calls.
            pass
        elif owner == "excel" and action == "pasteRange":
            self._powershell_setup(
                "seed-excel-clipboard",
                "$app=[System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application'); "
                "$app.ActiveWorkbook.ActiveSheet.Range('A1:B2').Copy() | Out-Null",
            )
        elif owner == "excel" and action == "copyRange":
            self._setup_action(
                "excel",
                "switchSheet",
                {"name": self.state["excel_main_sheet"]},
            )
        elif owner == "ppt" and action == "addAnimation":
            response, _ = self._setup_action(
                "ppt",
                "addShape",
                {"slideIndex": int(self.state["ppt_slide"]), "shapeType": "oval", "x": 650, "y": 80, "width": 80, "height": 80},
            )
            shape_id = _response_data(response).get("shapeId")
            if shape_id is not None:
                self.state["ppt_animation_shape_id"] = int(shape_id)
        elif owner == "word" and action == "generateTOC":
            self._setup_action("word", "applyStyle", {"range": 1, "style_name": "Heading 1"})

    def _capture_primary(self, owner: str, action: str, response: dict):
        data = _response_data(response)
        if owner == "excel":
            if action == "getContext" and data.get("currentSheet"):
                self.state["excel_main_sheet"] = data["currentSheet"]
            elif action == "createChart" and data.get("chartName"):
                self.state["excel_chart_name"] = data["chartName"]
            elif action == "createPivotTable" and data.get("tableName"):
                self.state["excel_pivot_name"] = data["tableName"]
            elif action == "saveAs":
                self.state["excel_main_name"] = Path(self.state["excel_final_path"]).name
            elif action == "convertFormat":
                self.state["excel_main_name"] = Path(self.state["excel_converted_path"]).name
        elif owner == "ppt":
            if action in {"addSlide", "duplicateSlide"}:
                if data.get("slideCount") is not None:
                    self.state["ppt_slide_count"] = int(data["slideCount"])
                if action == "duplicateSlide" and data.get("slideIndex") is not None:
                    self.state["ppt_duplicate_slide"] = int(data["slideIndex"])
            elif action == "insertSlidesFromFile" and data.get("slideCount") is not None:
                self.state["ppt_slide_count"] = int(data["slideCount"])
            elif action == "addTextBox" and data.get("shapeId") is not None:
                self.state["ppt_textbox_id"] = int(data["shapeId"])
            elif action == "addShape" and data.get("shapeId") is not None:
                self.state["ppt_shape_id"] = int(data["shapeId"])
            elif action == "duplicateShape" and data.get("shapeId") is not None:
                self.state["ppt_duplicate_shape_id"] = int(data["shapeId"])
            elif action == "insertPptImage" and data.get("shapeId") is not None:
                self.state["ppt_image_id"] = int(data["shapeId"])
            elif action == "insertImage" and data.get("shapeId") is not None:
                self.state["ppt_replace_image_id"] = int(data["shapeId"])
            elif action == "replacePptImage" and data.get("shapeId") is not None:
                self.state["ppt_replace_image_id"] = int(data["shapeId"])
            elif action == "insertPptTable" and data.get("shapeId") is not None:
                self.state["ppt_table_id"] = int(data["shapeId"])
            elif action == "insertPptChart" and data.get("shapeId") is not None:
                self.state["ppt_chart_id"] = int(data["shapeId"])

    def _after_primary(self, owner: str, action: str, response: dict):
        if not response.get("success"):
            return
        if owner == "excel" and action == "createWorkbook":
            self._setup_action("excel", "saveAs", {"filePath": self.state["excel_main_path"], "overwrite": True})
            self.state["excel_main_name"] = Path(self.state["excel_main_path"]).name
            self._setup_action("excel", "createWorkbook", {})
            self._setup_action("excel", "saveAs", {"filePath": self.state["excel_secondary_path"], "overwrite": True})
            self._setup_action("excel", "closeWorkbook", {"save": False})
            self._setup_action("excel", "switchWorkbook", {"name": self.state["excel_main_name"]})
        elif owner == "excel" and action == "openWorkbook":
            self._setup_action("excel", "switchWorkbook", {"name": self.state["excel_main_name"]})
            self._setup_action("excel", "closeWorkbook", {"name": Path(self.state["excel_secondary_path"]).name, "save": False})
        elif owner == "excel" and action == "setRangeData":
            fixtures = (
                ("E1:E2", [[1], [2]]),
                ("G1:H5", [["Name", "Value"], [" A ", 1], [" A ", 1], ["", None], [" B ", 2]]),
                ("I1:J5", [["Key", "Value"], ["A", 1], ["A", 1], ["B", 2], ["C", 3]]),
                ("K1:K3", [["a,b"], ["c,d"], ["e,f"]]),
                ("AJ1:AK2", [[1, 2], [3, 4]]),
                ("N1:O5", [["Group", "Amount"], ["A", 10], ["A", 20], ["B", 5], ["B", 7]]),
                ("Q1:Q1", [["EXCEL_FIND_TOKEN"]]),
                ("Z1:Z2", [["clear me"], ["clear me too"]]),
            )
            for cell_range, data in fixtures:
                self._setup_action("excel", "setRangeData", {"range": cell_range, "data": data})
        elif owner == "excel" and action == "protectWorkbook":
            self._setup_action("excel", "protectWorkbook", {"protect": False, "password": "wps-suite"})
        elif owner == "ppt" and action == "createPresentation":
            self._setup_action("ppt", "addSlide", {"layout": "title_content", "title": "WPS full suite"})
            self._setup_action("ppt", "saveAs", {"filePath": self.state["ppt_main_path"], "overwrite": True})
            self.state["ppt_main_name"] = Path(self.state["ppt_main_path"]).name
            self._setup_action("ppt", "createPresentation", {})
            self._setup_action("ppt", "addSlide", {"layout": "title", "title": "Imported fixture"})
            self._setup_action("ppt", "saveAs", {"filePath": self.state["ppt_secondary_path"], "overwrite": True})
            self._setup_action("ppt", "closePresentation", {"save": False})
            self._setup_action("ppt", "switchPresentation", {"name": self.state["ppt_main_name"]})
        elif owner == "ppt" and action == "openPresentation":
            self._setup_action("ppt", "closePresentation", {"save": False})
            self._setup_action("ppt", "switchPresentation", {"name": self.state["ppt_main_name"]})
        elif owner == "ppt" and action == "startSlideShow":
            self._powershell_setup(
                "close-ppt-slide-show",
                "$app=[System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwpp.Application'); "
                "if ($app.SlideShowWindows.Count -gt 0) {$app.SlideShowWindows.Item(1).View.Exit()}",
            )
        elif owner == "word" and action == "createDocument":
            self._setup_action("word", "saveAs", {"filePath": self.state["word_main_path"], "overwrite": True})
            self.state["word_main_name"] = Path(self.state["word_main_path"]).name
            self._setup_action("word", "openDocument", {"filePath": self.state["word_secondary_path"]})
            self._setup_action("word", "switchDocument", {"name": self.state["word_main_name"]})
        elif owner == "word" and action == "openDocument":
            self._setup_action("word", "switchDocument", {"name": self.state["word_main_name"]})

    def _record_skip(self, scenario: str, owner: str, action: str, params: dict):
        self.recorder.append({
            "scenario": scenario,
            "owner": owner,
            "action": action,
            "role": "primary",
            "status": "skip",
            "durationMs": 0,
            "params": params,
            "reason": "--skip-interactive",
        })
        print(f"[SKIP] {owner}.{action} (interactive)")

    def _record_case_exception(self, scenario: str, owner: str, action: str, params: dict, exc: Exception):
        detail = f"suite orchestration error: {type(exc).__name__}: {exc}"
        self.recorder.append({
            "scenario": scenario,
            "owner": owner,
            "action": action,
            "role": "primary",
            "status": "fail",
            "durationMs": 0,
            "params": params,
            "transportError": detail,
            "effectErrors": [],
            "response": None,
        })
        print(f"[FAIL] {owner}.{action} (suite orchestration error)")

    def run(self):
        validate_suite(self.catalog, self.state)
        _write_fixture_png(Path(self.state["fixture_image"]))
        _write_minimal_docx(Path(self.state["word_secondary_path"]), "WPS secondary fixture")
        _write_minimal_docx(Path(self.state["word_open_path"]), "WPS openDocument fixture")

        for scenario in SCENARIOS:
            print(f"\n=== {scenario.name} ({len(scenario.actions)} Actions) ===")
            for action in scenario.actions:
                params = build_action_params(scenario.owner, action, self.state)
                if self.skip_interactive and (scenario.owner, action) in INTERACTIVE_ACTIONS:
                    self._record_skip(scenario.name, scenario.owner, action, params)
                    continue
                record_start = len(self.recorder.records)
                try:
                    self._before_primary(scenario.owner, action)
                    # Rebuild because setup may have produced a shape/document identity.
                    params = build_action_params(scenario.owner, action, self.state)
                    self.catalog.validate_params(scenario.owner, action, params)
                    response, _ = self._invoke(
                        scenario.name,
                        scenario.owner,
                        action,
                        params,
                        "primary",
                    )
                    self._capture_primary(scenario.owner, action, response)
                    self._after_primary(scenario.owner, action, response)
                except Exception as exc:
                    new_records = self.recorder.records[record_start:]
                    has_primary = any(
                        record.get("role") == "primary"
                        and record.get("owner") == scenario.owner
                        and record.get("action") == action
                        for record in new_records
                    )
                    if has_primary:
                        self._record_harness(
                            f"after-{scenario.owner}-{action}",
                            False,
                            f"{type(exc).__name__}: {exc}",
                            0,
                        )
                    else:
                        self._record_case_exception(
                            scenario.name,
                            scenario.owner,
                            action,
                            params,
                            exc,
                        )

        return self._finish()

    def _finish(self):
        primary = [record for record in self.recorder.records if record.get("role") == "primary"]
        summary = {
            "declared": len(list(iter_declared_actions())),
            "recorded": len(primary),
            "passed": sum(record["status"] == "pass" for record in primary),
            "failed": sum(record["status"] == "fail" for record in primary),
            "skipped": sum(record["status"] == "skip" for record in primary),
            "setupFailed": sum(
                record.get("status") == "fail" and record.get("role") != "primary"
                for record in self.recorder.records
            ),
        }
        by_owner = {}
        for owner in ("bridge", "excel", "ppt", "word"):
            records = [record for record in primary if record["owner"] == owner]
            by_owner[owner] = {
                "total": len(records),
                "passed": sum(record["status"] == "pass" for record in records),
                "failed": sum(record["status"] == "fail" for record in records),
                "skipped": sum(record["status"] == "skip" for record in records),
            }
        report = {
            "suite": "Windows real-WPS full Action suite",
            "startedAt": self.started_at,
            "finishedAt": _utc_now(),
            "outputDirectory": str(self.output_dir),
            "summary": summary,
            "byOwner": by_owner,
            "records": self.recorder.records,
        }
        report_path = self.output_dir / "report.json"
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False),
            encoding="utf-8",
        )

        failures = [
            record for record in self.recorder.records
            if record.get("status") == "fail"
        ]
        failure_lines = []
        for record in failures:
            response = record.get("response") or {}
            failure_lines.extend((
                f"[{record.get('owner')}.{record.get('action')}] {record.get('scenario')}",
                f"role: {record.get('role')}",
                f"code: {response.get('code', '')}",
                f"error: {response.get('error', record.get('transportError', record.get('detail', '')))}",
                f"effectErrors: {record.get('effectErrors', [])}",
                f"traceId: {response.get('traceId', '')}",
                f"traceLog: {record.get('copiedTrace') or response.get('traceLog', '')}",
                "",
            ))
        if not failure_lines:
            failure_lines.append("No failures.\n")
        (self.output_dir / "failures.log").write_text("\n".join(failure_lines), encoding="utf-8")

        print("\n=== Full suite summary ===")
        print(json.dumps(summary, ensure_ascii=False))
        print(f"Report: {report_path}")
        print(f"Failures: {self.output_dir / 'failures.log'}")
        print("Note: Word has no closeDocument Action; close the generated Word documents after inspection.")
        return 1 if summary["failed"] or summary["setupFailed"] else 0


def _default_output_dir():
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return ROOT / "test-results" / f"wps-full-{stamp}"


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Result/artifact directory (default: test-results/wps-full-<timestamp>)",
    )
    parser.add_argument(
        "--skip-interactive",
        action="store_true",
        help="Skip PPT selection and slide-show Actions; all remain declared in coverage",
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Validate coverage and parameters without requiring Windows or WPS",
    )
    args = parser.parse_args(argv)

    output_dir = (args.output_dir or _default_output_dir()).resolve()
    catalog = ActionCatalog.from_path()
    state = default_suite_state(output_dir)
    count = validate_suite(catalog, state)
    if args.validate_only:
        print(f"Validated {count} Action cases: coverage and default parameters are valid.")
        return 0

    if sys.platform != "win32" or os.name != "nt":
        print("Real-WPS execution requires a Windows machine with WPS Office.", file=sys.stderr)
        return 2

    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Result directory: {output_dir}")
    print("Close or save unrelated WPS documents first; do not switch the active document during the run.")
    return LiveSuite(output_dir, skip_interactive=args.skip_interactive).run()


if __name__ == "__main__":
    raise SystemExit(main())
