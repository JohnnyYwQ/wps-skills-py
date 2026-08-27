import json
import io
import os
import queue
import tempfile
import threading
import unittest
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


class ControllerTraceTests(unittest.TestCase):
    def test_windows_controllers_launch_the_resolved_powershell_executable(self):
        cases = (
            (wps_excel, wps_excel.WpsExcelController),
            (wps_ppt, wps_ppt.WpsPptController),
            (wps_word, wps_word.WpsWordController),
        )
        executable = r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
        resolution = SimpleNamespace(
            available=True,
            powershell_executable=executable,
            selected_view_bitness=32,
            selected_registration=SimpleNamespace(clsid="{WPS-CLSID}"),
            diagnostic="",
        )
        for module, controller_type in cases:
            with self.subTest(controller=controller_type.__name__):
                process = Mock()
                process.stdin = io.StringIO()
                process.stdout = io.StringIO('{"ready":true}\n')
                process.stderr = io.StringIO("")
                process.pid = 4321
                process.poll.return_value = None
                process.wait.return_value = 0
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

    def test_retry_uses_new_req_id_but_keeps_same_action_trace(self):
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
                    controller._reinit_windows = Mock()
                    controller._read_result = Mock(
                        side_effect=[
                            {"success": False, "error": "RPC 服务器不可用"},
                            {"success": True, "data": {"ok": True}},
                        ]
                    )

                    result = controller._exec_windows(action, {}, trace=trace)

                    self.assertTrue(result["success"])
                    commands = [
                        json.loads(line)
                        for line in controller._ps_process.stdin.lines
                        if line.strip()
                    ]
                    self.assertEqual([1, 2], [command["reqId"] for command in commands])
                    self.assertEqual([1, 2], [command["attempt"] for command in commands])
                    self.assertEqual(
                        [trace.trace_id, trace.trace_id],
                        [command["traceId"] for command in commands],
                    )
                    events = [
                        json.loads(line)["event"]
                        for line in trace.log_path.read_text(encoding="utf-8").splitlines()
                    ]
                    self.assertIn("controller.retry.scheduled", events)

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
