import json
import os
import queue
import tempfile
import threading
import unittest
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
