import unittest
from unittest.mock import patch
from pathlib import Path
import sys
import json
import os
import tempfile
import io
import urllib.error
import urllib.request
from unittest.mock import MagicMock
from types import SimpleNamespace

import server
from action_trace import ActionTrace
from service_lifecycle import BridgeLifecycle, current_service_identity

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
            captured["headers"] = {
                key.lower(): value for key, value in request.header_items()
            }
            return _Response()

        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            trace = ActionTrace.start(component="call")
            service_health = {
                "status": "ok",
                **current_service_identity(instance_id="test-instance"),
            }
            with patch.object(call, "_urlopen", side_effect=_urlopen):
                result = call._post(
                    "findReplace",
                    {"find_text": "old", "replace_text": "new"},
                    app="word",
                    trace=trace,
                    service_health=service_health,
                )

        self.assertTrue(result["success"])
        self.assertEqual("word", captured["payload"]["app"])
        self.assertNotIn("app", captured["payload"]["params"])
        self.assertEqual(trace.trace_id, captured["payload"]["traceId"])
        self.assertEqual(trace.trace_id, result["traceId"])
        self.assertEqual(
            service_health["projectId"],
            captured["headers"]["x-wps-bridge-project-id"],
        )
        self.assertEqual(
            service_health["instanceId"],
            captured["headers"]["x-wps-bridge-instance-id"],
        )

    def test_http_identity_error_is_returned_as_json_not_connection_failure(self):
        service_health = {
            "status": "ok",
            **current_service_identity(instance_id="old-instance"),
        }
        body = io.BytesIO(json.dumps({
            "success": False,
            "code": "STALE_BRIDGE_INSTANCE",
            "error": "bridge 实例已变化",
        }).encode("utf-8"))
        error = urllib.error.HTTPError(
            call.BASE,
            409,
            "Conflict",
            {},
            body,
        )

        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ), patch.object(call, "_urlopen", side_effect=error):
            result = call._post(
                "setCellValue",
                {"cell": "A1", "value": 1},
                service_health=service_health,
            )

        self.assertFalse(result["success"])
        self.assertEqual("STALE_BRIDGE_INSTANCE", result["code"])

    def test_started_process_uses_kill_when_graceful_termination_times_out(self):
        process = MagicMock()
        process.poll.return_value = None
        process.wait.side_effect = [
            call.subprocess.TimeoutExpired("bridge/server.py", 3),
            None,
        ]

        call._terminate_started_process(process)

        process.terminate.assert_called_once_with()
        process.kill.assert_called_once_with()
        self.assertEqual(2, process.wait.call_count)

    def test_already_exited_started_process_is_not_terminated(self):
        process = MagicMock()
        process.poll.return_value = 1

        call._terminate_started_process(process)

        process.terminate.assert_not_called()
        process.kill.assert_not_called()

    def test_termination_error_still_falls_back_to_kill(self):
        process = MagicMock()
        process.poll.return_value = None
        process.terminate.side_effect = OSError("terminate failed")

        call._terminate_started_process(process)

        process.kill.assert_called_once_with()

    def test_poll_error_still_attempts_started_process_cleanup(self):
        process = MagicMock()
        process.poll.side_effect = OSError("poll failed")

        call._terminate_started_process(process)

        process.terminate.assert_called_once_with()


class CallServiceLifecycleRegressionTests(unittest.TestCase):
    def test_loopback_opener_does_not_inherit_environment_proxy(self):
        proxy_handlers = [
            handler
            for handler in call._LOOPBACK_OPENER.handlers
            if isinstance(handler, urllib.request.ProxyHandler)
        ]

        self.assertEqual([], proxy_handlers)

    def test_health_uses_loopback_urlopen_seam(self):
        class _Response:
            def __enter__(self):
                return self

            def __exit__(self, *args):
                return False

            def read(self):
                return b'{"status":"ok"}'

        with patch.object(call, "_urlopen", return_value=_Response()) as urlopen:
            health = call._health()

        self.assertEqual({"status": "ok"}, health)
        urlopen.assert_called_once_with(f"{call.BASE}/health", timeout=2)

    def test_health_timeout_is_unresponsive_not_absent(self):
        with patch.object(call, "_urlopen", side_effect=TimeoutError("busy")):
            probe = call._health(with_state=True)

        self.assertEqual("unresponsive", probe.state)
        self.assertIsNone(probe.health)

    def test_unresponsive_port_never_starts_a_second_bridge(self):
        probe = call.HealthProbe("unresponsive", error="timed out")
        with patch.object(call, "_health", return_value=probe), patch.object(
            call.subprocess,
            "Popen",
        ) as popen:
            result = call._ensure_server()

        self.assertFalse(result)
        self.assertEqual("BRIDGE_UNAVAILABLE", result.code)
        popen.assert_not_called()

    def test_non_object_health_is_rejected_without_trace_crash(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ), patch.object(call, "_health", return_value=[]), patch.object(
            call.subprocess,
            "Popen",
        ) as popen:
            result = call._ensure_server(trace=ActionTrace.start(component="test"))

        self.assertFalse(result)
        popen.assert_not_called()

    def test_json_null_health_response_is_not_treated_as_empty_port(self):
        probe = call.HealthProbe("responded", health=None)
        with patch.object(call, "_health", return_value=probe), patch.object(
            call.subprocess,
            "Popen",
        ) as popen:
            result = call._ensure_server()

        self.assertFalse(result)
        self.assertEqual("BRIDGE_INSTANCE_MISMATCH", result.code)
        popen.assert_not_called()

    def test_legacy_health_response_is_not_reused_as_current_checkout(self):
        with patch.object(call, "_health", return_value={"status": "ok"}), patch.object(
            call.subprocess,
            "Popen",
        ) as popen:
            result = call._ensure_server()

        self.assertFalse(result)
        popen.assert_not_called()

    def test_start_timeout_terminates_the_process_that_call_started(self):
        process = MagicMock()
        process.poll.return_value = None
        with patch.object(call, "_health", return_value=None), patch.object(
            call.os.path,
            "exists",
            return_value=True,
        ), patch.object(
            call,
            "server_log_path",
            return_value=(None, None),
        ), patch.object(
            call.subprocess,
            "Popen",
            return_value=process,
        ), patch.object(
            call.time,
            "sleep",
        ):
            result = call._ensure_server()

        self.assertFalse(result)
        process.terminate.assert_called_once_with()

    def test_start_poll_error_is_contained_and_process_is_reclaimed(self):
        process = MagicMock()
        process.poll.side_effect = OSError("poll failed")
        with patch.object(call, "_health", return_value=None), patch.object(
            call.os.path,
            "exists",
            return_value=True,
        ), patch.object(
            call,
            "server_log_path",
            return_value=(None, None),
        ), patch.object(
            call.subprocess,
            "Popen",
            return_value=process,
        ), patch.object(call.time, "sleep"):
            result = call._ensure_server()

        self.assertFalse(result)
        self.assertEqual("BRIDGE_START_TIMEOUT", result.code)
        process.terminate.assert_called_once_with()

    def test_windows_start_uses_hidden_process_group(self):
        expected = current_service_identity()
        ready_health = {
            "status": "ok",
            **current_service_identity(instance_id="new-instance"),
        }
        process = MagicMock()

        class _StartupInfo:
            def __init__(self):
                self.dwFlags = 0

        with patch.object(call, "_health", side_effect=[None, ready_health]), patch.object(
            call,
            "current_service_identity",
            return_value=expected,
        ), patch.object(
            call.os.path,
            "exists",
            return_value=True,
        ), patch.object(
            call,
            "server_log_path",
            return_value=(None, None),
        ), patch.object(
            call.os,
            "name",
            "nt",
        ), patch.object(
            call.subprocess,
            "STARTUPINFO",
            _StartupInfo,
            create=True,
        ), patch.object(
            call.subprocess,
            "STARTF_USESHOWWINDOW",
            0x01,
            create=True,
        ), patch.object(
            call.subprocess,
            "CREATE_NEW_PROCESS_GROUP",
            0x200,
            create=True,
        ), patch.object(
            call.subprocess,
            "Popen",
            return_value=process,
        ) as popen, patch.object(call.time, "sleep"):
            result = call._ensure_server()

        self.assertTrue(result)
        self.assertTrue(result.started)
        self.assertEqual(expected["projectId"], result.health["projectId"])
        kwargs = popen.call_args.kwargs
        self.assertEqual(0x200, kwargs["creationflags"])
        self.assertEqual(0x01, kwargs["startupinfo"].dwFlags)
        self.assertNotIn("start_new_session", kwargs)


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
            identity = current_service_identity(instance_id="test-instance")
            handler.headers.update({
                "X-WPS-Bridge-Project-Id": identity["projectId"],
                "X-WPS-Bridge-Instance-Id": identity["instanceId"],
            })
            handler.server = SimpleNamespace(
                lifecycle=BridgeLifecycle(identity, idle_timeout_seconds=60)
            )
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
            identity = current_service_identity(instance_id="test-instance")
            handler.headers = {
                "Content-Length": "1",
                "X-WPS-Bridge-Project-Id": identity["projectId"],
                "X-WPS-Bridge-Instance-Id": identity["instanceId"],
            }
            handler.server = SimpleNamespace(
                lifecycle=BridgeLifecycle(identity, idle_timeout_seconds=60)
            )
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
