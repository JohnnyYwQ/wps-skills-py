import json
import io
import os
import queue
import tempfile
import threading
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from action_trace import ActionTrace
import wps_excel
import wps_ppt
import wps_word


class _InputCapture:
    def __init__(self):
        self.lines = []

    def write(self, value):
        self.lines.append(value)

    def flush(self):
        pass


class _RunningProcess:
    def __init__(self):
        self.stdin = _InputCapture()
        self.pid = 4321

    def poll(self):
        return None


def _windows_com_resolution(executable, bitness=64):
    return SimpleNamespace(
        available=True,
        powershell_executable=executable,
        selected_view_bitness=bitness,
        selected_registration=SimpleNamespace(clsid="{WPS-CLSID}"),
        diagnostic="",
    )


def _ready_powershell_process(stdin=None):
    process = Mock()
    process.stdin = io.StringIO() if stdin is None else stdin
    process.stdout = io.StringIO('{"ready":true}\n')
    process.stderr = io.StringIO("")
    process.pid = 4321
    process.poll.return_value = None
    process.wait.return_value = 0
    return process


class ControllerTraceTests(unittest.TestCase):
    def test_excel_fake_process_receives_attach_first_startup_script(self):
        resolution = _windows_com_resolution(
            r"C:\\Windows\\System32\\powershell.exe",
        )
        process = _ready_powershell_process(_InputCapture())

        with patch.object(wps_excel, "IS_WINDOWS", True), patch.object(
            wps_excel, "IS_LINUX", False,
        ), patch.object(
            wps_excel, "resolve_com_runtime", return_value=resolution,
        ), patch.object(wps_excel.subprocess, "Popen", return_value=process) as popen:
            controller = wps_excel.WpsExcelController()
            try:
                script = Path(popen.call_args.args[0][-1]).read_text(encoding="utf-8-sig")
            finally:
                controller.close()

        self.assertLess(
            script.index("GetActiveObject('Ket.Application')"),
            script.index("New-Object -ComObject 'Ket.Application'"),
        )
        self.assertNotIn("Workbooks.Add", script[:script.index("function Exec-ping")])
        self.assertIn("EXIT\n", process.stdin.lines)

    def test_excel_command_derives_active_workbook_requirement_from_contract(self):
        controller = object.__new__(wps_excel.WpsExcelController)
        controller._ps_process = _RunningProcess()
        controller._ready = True
        controller._id_counter = 0
        controller._id_lock = threading.Lock()
        controller._stop = None
        controller._read_result = Mock(return_value={"success": True})

        controller._exec_windows("setCellValue", {"row": 1, "col": 1, "value": 42})
        controller._exec_windows("createWorkbook", {})

        commands = [
            json.loads(line)
            for line in controller._ps_process.stdin.lines
            if line.strip()
        ]
        self.assertEqual([True, False], [
            command["requiresActiveWorkbook"] for command in commands
        ])

    def test_ppt_fake_process_receives_attach_first_startup_script(self):
        resolution = _windows_com_resolution(
            r"C:\\Windows\\System32\\powershell.exe",
        )
        process = _ready_powershell_process(_InputCapture())

        with patch.object(wps_ppt, "IS_WINDOWS", True), patch.object(
            wps_ppt, "IS_LINUX", False,
        ), patch.object(
            wps_ppt, "resolve_com_runtime", return_value=resolution,
        ), patch.object(wps_ppt.subprocess, "Popen", return_value=process) as popen:
            controller = wps_ppt.WpsPptController()
            try:
                script = Path(popen.call_args.args[0][-1]).read_text(encoding="utf-8-sig")
            finally:
                controller.close()

        self.assertLess(
            script.index("GetActiveObject('Kwpp.Application')"),
            script.index("New-Object -ComObject 'Kwpp.Application'"),
        )
        self.assertNotIn("Presentations.Add", script[:script.index("function Exec-ping")])
        self.assertIn("EXIT\n", process.stdin.lines)

    def test_ppt_command_derives_active_presentation_requirement_from_contract(self):
        controller = object.__new__(wps_ppt.WpsPptController)
        controller._ps_process = _RunningProcess()
        controller._ready = True
        controller._id_counter = 0
        controller._id_lock = threading.Lock()
        controller._stop = None
        controller._read_result = Mock(return_value={"success": True})

        controller._exec_windows("addSlide", {})
        controller._exec_windows("createPresentation", {})

        commands = [
            json.loads(line)
            for line in controller._ps_process.stdin.lines
            if line.strip()
        ]
        self.assertEqual([True, False], [
            command["requiresActivePresentation"] for command in commands
        ])

    def test_word_fake_process_receives_attach_first_startup_script(self):
        resolution = _windows_com_resolution(
            r"C:\\Windows\\System32\\powershell.exe",
        )
        process = _ready_powershell_process(_InputCapture())

        with patch.object(wps_word, "IS_WINDOWS", True), patch.object(
            wps_word, "IS_LINUX", False,
        ), patch.object(
            wps_word, "resolve_com_runtime", return_value=resolution,
        ), patch.object(wps_word.subprocess, "Popen", return_value=process) as popen:
            controller = wps_word.WpsWordController()
            try:
                script = Path(popen.call_args.args[0][-1]).read_text(encoding="utf-8-sig")
            finally:
                controller.close()

        self.assertLess(
            script.index("GetActiveObject('Kwps.Application')"),
            script.index("New-Object -ComObject 'Kwps.Application'"),
        )
        self.assertNotIn("Documents.Add", script[:script.index("function Exec-ping")])
        self.assertIn("EXIT\n", process.stdin.lines)

    def test_word_command_derives_active_document_requirement_from_contract(self):
        controller = object.__new__(wps_word.WpsWordController)
        controller._ps_process = _RunningProcess()
        controller._ready = True
        controller._id_counter = 0
        controller._id_lock = threading.Lock()
        controller._stop = None
        controller._read_result = Mock(return_value={"success": True})

        controller._exec_windows("getDocumentText", {})
        controller._exec_windows("createDocument", {})
        controller._exec_windows("openDocument", {"filePath": r"C:\\tmp\\report.docx"})

        commands = [
            json.loads(line)
            for line in controller._ps_process.stdin.lines
            if line.strip()
        ]
        self.assertEqual([True, False, False], [
            command["requiresActiveDocument"] for command in commands
        ])

    def test_windows_controllers_launch_the_resolved_powershell_executable(self):
        cases = (
            (wps_excel, wps_excel.WpsExcelController),
            (wps_ppt, wps_ppt.WpsPptController),
            (wps_word, wps_word.WpsWordController),
        )
        executable = r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
        resolution = _windows_com_resolution(executable, bitness=32)
        for module, controller_type in cases:
            with self.subTest(controller=controller_type.__name__):
                process = _ready_powershell_process()
                with patch.object(module, "IS_WINDOWS", True), patch.object(
                    module,
                    "IS_LINUX",
                    False,
                ), patch.object(
                    module,
                    "resolve_com_runtime",
                    return_value=resolution,
                ), patch.object(
                    module.subprocess,
                    "Popen",
                    return_value=process,
                ) as popen:
                    controller = controller_type()
                    try:
                        command = popen.call_args.args[0]
                        self.assertEqual(executable, command[0])
                        self.assertEqual("replace", popen.call_args.kwargs["errors"])
                    finally:
                        controller.close()

    def test_windows_controller_startup_failure_includes_exit_code_and_stderr(self):
        cases = (
            (wps_excel, wps_excel.WpsExcelController),
            (wps_ppt, wps_ppt.WpsPptController),
            (wps_word, wps_word.WpsWordController),
        )
        resolution = SimpleNamespace(
            available=True,
            powershell_executable=r"C:\Windows\System32\powershell.exe",
            selected_view_bitness=64,
            selected_registration=SimpleNamespace(clsid="{WPS-CLSID}"),
            diagnostic="",
        )
        for module, controller_type in cases:
            with self.subTest(controller=controller_type.__name__):
                process = Mock()
                process.stdin = io.StringIO()
                process.stdout = io.StringIO("")
                process.stderr = io.StringIO("native COM activation failure\n")
                process.pid = 4321
                process.poll.return_value = 7
                process.wait.return_value = 7
                with patch.object(module, "IS_WINDOWS", True), patch.object(
                    module,
                    "IS_LINUX",
                    False,
                ), patch.object(
                    module,
                    "resolve_com_runtime",
                    return_value=resolution,
                ), patch.object(
                    module.subprocess,
                    "Popen",
                    return_value=process,
                ):
                    with self.assertRaises(RuntimeError) as raised:
                        controller_type()

                message = str(raised.exception)
                self.assertIn("exit code 7", message)
                self.assertIn("native COM activation failure", message)

    def test_controller_close_waits_for_clean_powershell_exit(self):
        controller_types = (
            wps_excel.WpsExcelController,
            wps_ppt.WpsPptController,
            wps_word.WpsWordController,
        )
        for controller_type in controller_types:
            with self.subTest(controller=controller_type.__name__):
                process = Mock()
                process.stdin = _InputCapture()
                process.poll.return_value = None
                process.wait.return_value = 0
                controller = object.__new__(controller_type)
                controller._ps_process = process
                controller._ready = True
                controller._stop = threading.Event()

                controller.close()

                self.assertIn("EXIT\n", process.stdin.lines)
                process.wait.assert_called_once()
                process.kill.assert_not_called()
                self.assertFalse(controller._ready)
                self.assertIsNone(controller._ps_process)

    def test_windows_controllers_make_one_attempt_even_after_rpc_disconnect(self):
        cases = (
            (wps_excel.WpsExcelController, "setCellValue"),
            (wps_ppt.WpsPptController, "addSlide"),
            (wps_word.WpsWordController, "insertText"),
        )
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            for controller_type, action in cases:
                with self.subTest(controller=controller_type.__name__):
                    trace = ActionTrace.start(component="server")
                    controller = object.__new__(controller_type)
                    controller._ps_process = _RunningProcess()
                    controller._ready = True
                    controller._id_counter = 0
                    controller._id_lock = threading.Lock()
                    controller._stop = None
                    controller._read_result = Mock(
                        return_value={
                            "success": False,
                            "error": "RPC 服务器不可用",
                        },
                    )

                    result = controller._exec_windows(action, {}, trace=trace)

                    self.assertFalse(result["success"])
                    commands = [
                        json.loads(line)
                        for line in controller._ps_process.stdin.lines
                        if line.strip()
                    ]
                    self.assertEqual([1], [command["reqId"] for command in commands])
                    self.assertEqual([1], [command["attempt"] for command in commands])
                    self.assertEqual(
                        [trace.trace_id],
                        [command["traceId"] for command in commands],
                    )
                    events = [
                        json.loads(line)["event"]
                        for line in trace.log_path.read_text(encoding="utf-8").splitlines()
                    ]
                    self.assertNotIn("controller.retry.scheduled", events)

    def test_windows_controllers_use_the_runtime_correlation_id(self):
        cases = (
            (wps_excel.WpsExcelController, "setCellValue"),
            (wps_ppt.WpsPptController, "addSlide"),
            (wps_word.WpsWordController, "insertText"),
        )
        for controller_type, action in cases:
            with self.subTest(controller=controller_type.__name__):
                controller = object.__new__(controller_type)
                controller._ps_process = _RunningProcess()
                controller._ready = True
                controller._id_counter = 0
                controller._id_lock = threading.Lock()
                controller._stop = None
                controller._read_result = Mock(return_value={"success": True})

                controller._exec_windows(
                    action,
                    {},
                    correlation_id="runtime-retry-2",
                )

                command = json.loads(controller._ps_process.stdin.lines[0])
                self.assertEqual("runtime-retry-2", command["reqId"])

    def test_windows_controller_timeout_is_structured(self):
        cases = (
            (wps_excel, wps_excel.WpsExcelController),
            (wps_ppt, wps_ppt.WpsPptController),
            (wps_word, wps_word.WpsWordController),
        )
        for module, controller_type in cases:
            with self.subTest(controller=controller_type.__name__):
                controller = object.__new__(controller_type)
                controller._ready = True
                controller._ps_process = None
                controller._stop = None
                controller._stderr_queue = queue.Queue()

                with patch.object(module, "bounded_timeout", return_value=0):
                    result = controller._read_result(7)

                self.assertEqual("ACTION_EXECUTION_TIMEOUT", result["code"])

    def test_backend_timing_and_stderr_are_recorded_but_not_returned(self):
        cases = (
            (wps_excel.WpsExcelController, "excel"),
            (wps_ppt.WpsPptController, "ppt"),
            (wps_word.WpsWordController, "word"),
        )
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            for controller_type, app in cases:
                with self.subTest(controller=controller_type.__name__):
                    trace = ActionTrace.start(component="server")
                    controller = object.__new__(controller_type)
                    controller._stop = None
                    controller._ps_process = None
                    controller._queue = queue.Queue()
                    controller._stderr_queue = queue.Queue()
                    controller._queue.put(json.dumps({
                        "success": True,
                        "data": {"ok": True},
                        "reqId": 7,
                        "_trace": {
                            "traceId": trace.trace_id,
                            "backendElapsedMs": 12.5,
                            "attempt": 1,
                        },
                    }))
                    controller._stderr_queue.put("native diagnostic\n")

                    result = controller._read_result(7, trace=trace, attempt=1)

                    self.assertTrue(result["success"])
                    self.assertNotIn("_trace", result)
                    rows = [
                        json.loads(line)
                        for line in trace.log_path.read_text(encoding="utf-8").splitlines()
                    ]
                    stderr_row = next(row for row in rows if row["event"] == "powershell.stderr")
                    response_row = next(
                        row for row in rows if row["event"] == "powershell.response.received"
                    )
                    self.assertEqual(app, stderr_row["app"])
                    self.assertEqual("native diagnostic", stderr_row["error"])
                    self.assertEqual(trace.trace_id, response_row["backendTraceId"])
                    self.assertEqual(1, response_row["backendAttempt"])
                    self.assertEqual(12.5, response_row["backendElapsedMs"])


if __name__ == "__main__":
    unittest.main()
