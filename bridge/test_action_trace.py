import json
import os
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch

import action_trace
from action_trace import ActionTrace


class ActionTraceTests(unittest.TestCase):
    def test_trace_writes_jsonl_and_decorates_every_result(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp, "WPS_TRACE": "info"},
            clear=False,
        ):
            trace = ActionTrace.start(component="call")
            trace.event("call.started", action="addSlide")
            result = trace.decorate({"success": True, "data": {"slideIndex": 1}})

            self.assertTrue(result["success"])
            self.assertEqual(trace.trace_id, result["traceId"])
            self.assertEqual(str(trace.log_path), result["traceLog"])
            self.assertTrue(trace.log_path.is_file())

            rows = [json.loads(line) for line in trace.log_path.read_text(encoding="utf-8").splitlines()]
            self.assertEqual("call.started", rows[-1]["event"])
            self.assertEqual("call", rows[-1]["component"])
            self.assertEqual(trace.trace_id, rows[-1]["traceId"])

    def test_resume_appends_to_the_same_action_file(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            client = ActionTrace.start(component="call")
            client.event("http.sent")
            server = ActionTrace.resume(client.trace_id, component="server")
            server.event("http.received")

            self.assertEqual(client.log_path, server.log_path)
            rows = [json.loads(line) for line in client.log_path.read_text(encoding="utf-8").splitlines()]
            self.assertEqual(["http.sent", "http.received"], [row["event"] for row in rows])

    def test_info_mode_omits_debug_payload(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp, "WPS_TRACE": "info"},
            clear=False,
        ):
            trace = ActionTrace.start(component="call")

            self.assertEqual({}, trace.debug_fields(params={"text": "private"}))

    def test_debug_mode_redacts_secrets_and_summarizes_document_content(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp, "WPS_TRACE": "debug"},
            clear=False,
        ):
            trace = ActionTrace.start(component="call")
            fields = trace.debug_fields(
                params={
                    "token": "secret-token",
                    "text": "confidential document text",
                    "filePath": "C:/work/report.docx",
                    "slideIndex": 2,
                }
            )

            params = fields["params"]
            self.assertEqual("<redacted>", params["token"])
            self.assertEqual("str", params["text"]["type"])
            self.assertEqual(len("confidential document text"), params["text"]["length"])
            self.assertNotIn("confidential", json.dumps(params["text"]))
            self.assertEqual("C:/work/report.docx", params["filePath"])
            self.assertEqual(2, params["slideIndex"])

    def test_start_removes_trace_files_older_than_one_day(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            traces_dir = Path(tmp) / "traces" / "2000-01-01"
            traces_dir.mkdir(parents=True)
            old_file = traces_dir / "act-20000101-deadbeef.jsonl"
            old_file.write_text("{}\n", encoding="utf-8")
            old_time = time.time() - 90000
            os.utime(old_file, (old_time, old_time))

            recent_dir = Path(tmp) / "traces" / "2099-01-01"
            recent_dir.mkdir(parents=True)
            recent_file = recent_dir / "recent.jsonl"
            recent_file.write_text("{}\n", encoding="utf-8")

            old_server_log = Path(tmp) / "server-2000-01-01.log"
            old_server_log.write_text("old server output\n", encoding="utf-8")
            os.utime(old_server_log, (old_time, old_time))

            unrelated_file = Path(tmp) / "keep-me.txt"
            unrelated_file.write_text("not a trace log\n", encoding="utf-8")
            os.utime(unrelated_file, (old_time, old_time))

            ActionTrace.start(component="call")

            self.assertFalse(old_file.exists())
            self.assertFalse(old_server_log.exists())
            self.assertTrue(recent_file.exists())
            self.assertTrue(unrelated_file.exists())

    def test_cleanup_failure_warns_but_does_not_block_action(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ), patch.object(
            action_trace,
            "_remove_old_logs_unchecked",
            side_effect=OSError("cleanup denied"),
        ):
            trace = ActionTrace.start(component="call")
            trace.event("action.started", action="addSlide")
            result = trace.decorate({"success": True})

            self.assertTrue(result["success"])
            self.assertIn("过期日志清理失败", result["traceWarning"])
            self.assertTrue(trace.log_path.is_file())

    def test_no_writable_directory_still_returns_stable_trace_fields(self):
        with patch.object(
            action_trace,
            "_select_log_root",
            return_value=(None, "no writable directory"),
        ):
            trace = ActionTrace.start(component="call")
            result = trace.decorate({"success": False, "error": "failed"})

            self.assertTrue(result["traceId"].startswith("act-"))
            self.assertIsNone(result["traceLog"])
            self.assertEqual("no writable directory", result["traceWarning"])


if __name__ == "__main__":
    unittest.main()
