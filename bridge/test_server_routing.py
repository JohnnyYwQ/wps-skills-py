import unittest
from unittest.mock import patch
from pathlib import Path
import sys
import json
import os
import tempfile
import io

import server
from action_trace import ActionTrace

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))
import call


class _RecordingController:
    def __init__(self, app):
        self.app = app

    def execute(self, action, params, trace=None):
        return {
            "success": True,
            "data": {
                "routedApp": self.app,
                "action": action,
                "params": params,
            },
        }


class ActionRoutingTests(unittest.TestCase):
    def _dispatch(self, action, params=None, app=None):
        with patch.object(
            server,
            "get_app_controller",
            side_effect=lambda selected, trace=None: _RecordingController(selected),
        ):
            if app is None:
                return server.dispatch(action, params or {})
            return server.dispatch(action, params or {}, app=app)

    def test_explicit_app_routes_ambiguous_action_to_word(self):
        result = self._dispatch(
            "findReplace",
            {"find_text": "old", "replace_text": "new"},
            app="word",
        )

        self.assertTrue(result["success"])
        self.assertEqual("word", result["data"]["routedApp"])

    def test_params_app_is_supported_but_not_forwarded_to_controller(self):
        result = self._dispatch(
            "insertImage",
            {"app": "word", "imagePath": "C:/tmp/image.png"},
        )

        self.assertTrue(result["success"])
        self.assertEqual("word", result["data"]["routedApp"])
        self.assertNotIn("app", result["data"]["params"])

    def test_ambiguous_action_without_app_fails_instead_of_guessing(self):
        result = self._dispatch("findReplace", {"find": "old", "replace": "new"})

        self.assertFalse(result["success"])
        self.assertEqual("AMBIGUOUS_ACTION", result["code"])
        self.assertEqual(["excel", "word"], result["supportedApps"])

    def test_unique_action_keeps_automatic_routing(self):
        result = self._dispatch("setCellValue", {"cell": "A1", "value": 1})

        self.assertTrue(result["success"])
        self.assertEqual("excel", result["data"]["routedApp"])

    def test_explicit_wrong_app_is_rejected_before_controller(self):
        result = self._dispatch("setCellValue", {"cell": "A1"}, app="word")

        self.assertFalse(result["success"])
        self.assertEqual("ACTION_NOT_SUPPORTED_FOR_APP", result["code"])
        self.assertEqual(["excel"], result["supportedApps"])

    def test_server_actions_are_listed_once_as_common(self):
        actions = server.get_action_list()

        self.assertEqual(
            [{"action": "ping", "app": "common"}],
            [item for item in actions if item["action"] == "ping"],
        )
        self.assertEqual(
            [{"action": "wireCheck", "app": "common"}],
            [item for item in actions if item["action"] == "wireCheck"],
        )


class CallArgumentTests(unittest.TestCase):
    def test_app_option_is_parsed_separately_from_action_params(self):
        action, params, app = call._load_params_from_args(
            ["findReplace", "--app", "word", '{"find_text":"old"}']
        )

        self.assertEqual("findReplace", action)
        self.assertEqual({"find_text": "old"}, params)
        self.assertEqual("word", app)

    def test_post_sends_app_as_top_level_routing_field(self):
        captured = {}

        class _Response:
            def __enter__(self):
                return self

            def __exit__(self, *args):
                return False

            def read(self):
                return b'{"success":true}'

        def _urlopen(request, timeout):
            captured["payload"] = json.loads(request.data.decode("utf-8"))
            return _Response()

        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            trace = ActionTrace.start(component="call")
            with patch.object(call.urllib.request, "urlopen", side_effect=_urlopen):
                result = call._post(
                    "findReplace",
                    {"find_text": "old", "replace_text": "new"},
                    app="word",
                    trace=trace,
                )

        self.assertTrue(result["success"])
        self.assertEqual("word", captured["payload"]["app"])
        self.assertNotIn("app", captured["payload"]["params"])
        self.assertEqual(trace.trace_id, captured["payload"]["traceId"])
        self.assertEqual(trace.trace_id, result["traceId"])


class HttpTraceTests(unittest.TestCase):
    def test_dispatch_response_resumes_and_returns_client_trace(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            client_trace = ActionTrace.start(component="test-client")
            payload = json.dumps(
                {
                    "traceId": client_trace.trace_id,
                    "action": "setCellValue",
                    "params": {"cell": "A1", "value": 1},
                }
            ).encode("utf-8")
            handler = object.__new__(server.Handler)
            handler.path = "/dispatch"
            handler.headers = {
                "Content-Length": str(len(payload)),
                "X-WPS-Trace-Id": client_trace.trace_id,
            }
            handler.rfile = io.BytesIO(payload)
            captured = {}
            handler._send = lambda obj, code=200: captured.update(result=obj, code=code)

            with patch.object(
                server,
                "get_app_controller",
                side_effect=lambda selected, trace=None: _RecordingController(selected),
            ):
                handler.do_POST()

            result = captured["result"]
            self.assertTrue(result["success"])
            self.assertEqual(client_trace.trace_id, result["traceId"])
            self.assertEqual(str(client_trace.log_path), result["traceLog"])
            events = [
                json.loads(line)["event"]
                for line in client_trace.log_path.read_text(encoding="utf-8").splitlines()
            ]
            self.assertIn("http.request.received", events)
            self.assertIn("route.selected", events)
            self.assertIn("http.response.ready", events)

    def test_invalid_json_response_still_returns_trace_location(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            handler = object.__new__(server.Handler)
            handler.path = "/dispatch"
            handler.headers = {"Content-Length": "1"}
            handler.rfile = io.BytesIO(b"{")
            captured = {}
            handler._send = lambda obj, code=200: captured.update(result=obj, code=code)

            handler.do_POST()

            result = captured["result"]
            self.assertEqual(400, captured["code"])
            self.assertFalse(result["success"])
            self.assertEqual("INVALID_JSON", result["code"])
            self.assertTrue(result["traceId"].startswith("act-"))
            self.assertTrue(Path(result["traceLog"]).is_file())


if __name__ == "__main__":
    unittest.main()
