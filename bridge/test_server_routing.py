import unittest
from unittest.mock import patch
from pathlib import Path
import sys
import json
import os
import errno
import tempfile
import io
import urllib.error
import urllib.request
from contextlib import nullcontext
from unittest.mock import MagicMock
from types import SimpleNamespace

import server
from action_trace import ActionTrace
from service_lifecycle import (
    BridgeLifecycle,
    ServiceStartResult,
    build_service_identity,
    current_service_identity,
)

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
    def setUp(self):
        listener_patcher = patch.object(
            call,
            "_listener_identity",
            return_value={"pid": None, "createdAt": None},
        )
        listener_patcher.start()
        self.addCleanup(listener_patcher.stop)

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

    def test_health_trace_records_client_request_and_timeout(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            trace = ActionTrace.start(component="call")
            with patch.object(call, "_urlopen", side_effect=TimeoutError("busy")):
                probe = call._health(with_state=True, trace=trace)

            events = [
                json.loads(line)
                for line in trace.log_path.read_text(encoding="utf-8").splitlines()
            ]

        self.assertEqual("unresponsive", probe.state)
        self.assertEqual(
            ["bridge.health.requested", "bridge.health.failed"],
            [row["event"] for row in events],
        )
        self.assertEqual("TimeoutError", events[-1]["errorType"])
        self.assertGreaterEqual(events[-1]["elapsedMs"], 0)

    def test_health_trace_correlates_client_request_and_response(self):
        class _Response:
            status = 200

            def __enter__(self):
                return self

            def __exit__(self, *args):
                return False

            def read(self):
                return b'{"status":"ok"}'

        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            trace = ActionTrace.start(component="call")
            with patch.object(call, "_urlopen", return_value=_Response()) as urlopen:
                probe = call._health(with_state=True, trace=trace)

            events = [
                json.loads(line)
                for line in trace.log_path.read_text(encoding="utf-8").splitlines()
            ]

        request = urlopen.call_args.args[0]
        self.assertEqual(trace.trace_id, request.get_header("X-wps-trace-id"))
        self.assertEqual("responded", probe.state)
        self.assertEqual(
            ["bridge.health.requested", "bridge.health.responded"],
            [row["event"] for row in events],
        )
        self.assertEqual(200, events[-1]["httpStatus"])
        self.assertEqual(15, events[-1]["bodyBytes"])

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

    def test_health_timeout_reports_listener_identity_before_and_after_probe(self):
        probe = call.HealthProbe("unresponsive", error="timed out")
        listener_before = {"pid": 111, "createdAt": "before-time"}
        listener_after = {"pid": 111, "createdAt": "before-time"}
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp},
            clear=False,
        ):
            trace = ActionTrace.start(component="test")
            with patch.object(call, "_health", return_value=probe), patch.object(
                call,
                "_listener_identity",
                side_effect=[listener_before, listener_after],
                create=True,
            ), patch.object(call.subprocess, "Popen") as popen:
                result = call._ensure_server(trace=trace)

            events = [
                json.loads(line)
                for line in trace.log_path.read_text(encoding="utf-8").splitlines()
            ]

        unavailable = next(
            row for row in events if row["event"] == "bridge.health.unavailable"
        )
        self.assertFalse(result)
        self.assertEqual("unknown_owner", result.disposition)
        self.assertEqual(listener_before, result.listener_before)
        self.assertEqual(listener_after, result.listener_after)
        self.assertEqual(listener_before, unavailable["listenerBefore"])
        self.assertEqual(listener_after, unavailable["listenerAfter"])
        popen.assert_not_called()

    def test_listener_without_confirmed_health_is_unknown_and_never_replaced(self):
        listener = {"pid": 222, "createdAt": "listener-start"}
        with patch.object(
            call,
            "_health",
            return_value=call.HealthProbe("absent", error="connection refused"),
        ), patch.object(
            call,
            "_listener_identity",
            return_value=listener,
        ), patch.object(call.subprocess, "Popen") as popen:
            result = call._ensure_server()

        self.assertFalse(result)
        self.assertEqual("unknown_owner", result.disposition)
        self.assertEqual(listener, result.listener_before)
        self.assertEqual(listener, result.listener_after)
        popen.assert_not_called()

    def test_current_checkout_health_without_launch_identity_is_unknown(self):
        incomplete = {
            "status": "ok",
            **current_service_identity(instance_id="incomplete-instance"),
        }
        with patch.object(call, "_health", return_value=incomplete), patch.object(
            call.subprocess,
            "Popen",
        ) as popen:
            result = call._ensure_server()

        self.assertFalse(result)
        self.assertEqual("unknown_owner", result.disposition)
        popen.assert_not_called()

    def test_other_checkout_health_is_classified_as_foreign(self):
        foreign_health = {
            "status": "ok",
            **build_service_identity(
                "/workspace/other",
                "foreign-code",
                instance_id="foreign-instance",
            ),
        }
        with patch.object(call, "_health", return_value=foreign_health), patch.object(
            call.subprocess,
            "Popen",
        ) as popen:
            result = call._ensure_server()

        self.assertFalse(result)
        self.assertEqual("foreign", result.disposition)
        popen.assert_not_called()

    def test_reuse_accepts_current_health_when_os_listener_is_unobservable(self):
        health = {
            "status": "ok",
            **current_service_identity(instance_id="claimed-instance"),
            "launchId": "claimed-launch",
            "launchedByPid": 123,
            "serverPath": call._server_path(),
            "launchStartedAt": "2026-08-27T01:02:03.004Z",
            "pid": 124,
            "serverPid": 124,
        }
        unobservable_listener = {"pid": None, "createdAt": None}
        with patch.object(call, "_health", return_value=health), patch.object(
            call,
            "_listener_identity",
            return_value=unobservable_listener,
        ):
            result = call._ensure_server()

        self.assertTrue(result)
        self.assertFalse(result.started)
        self.assertEqual("reused", result.disposition)
        self.assertEqual("claimed-instance", result.health["instanceId"])

    def test_start_winner_reprobes_inside_lock_and_reuses_ready_bridge(self):
        ready_health = {
            "status": "ok",
            **current_service_identity(instance_id="winner-instance"),
            "launchId": "winner-launch",
            "launchedByPid": 123,
            "serverPath": call._server_path(),
            "launchStartedAt": "2026-08-27T01:02:03.004Z",
            "pid": 124,
            "serverPid": 124,
        }
        with patch.object(
            call,
            "_health",
            side_effect=[None, ready_health],
        ), patch.object(
            call,
            "_startup_lock",
            return_value=nullcontext(),
            create=True,
        ), patch.object(
            call,
            "_listener_identity",
            side_effect=[
                {"pid": None, "createdAt": None},
                {"pid": None, "createdAt": None},
                {"pid": 124, "createdAt": "winner-time"},
                {"pid": 124, "createdAt": "winner-time"},
            ],
        ), patch.object(call.subprocess, "Popen") as popen:
            result = call._ensure_server()

        self.assertTrue(result)
        self.assertFalse(result.started)
        self.assertEqual("reused", result.disposition)
        self.assertEqual("winner-instance", result.health["instanceId"])
        popen.assert_not_called()

    def test_self_started_uses_health_identity_when_listener_is_unobservable(self):
        process = MagicMock()
        process.pid = 456
        process.poll.return_value = None
        ready_health = {
            "status": "ok",
            **current_service_identity(instance_id="new-instance"),
            "launchId": "launch-a",
            "pid": 456,
            "serverPid": 456,
            "launchedByPid": os.getpid(),
            "serverPath": call._server_path(),
            "launchStartedAt": "2026-08-27T01:02:03.004Z",
        }
        with patch.object(
            call,
            "_health",
            side_effect=[None, None, ready_health],
        ), patch.object(call.os.path, "exists", return_value=True), patch.object(
            call,
            "server_log_path",
            return_value=(None, None),
        ), patch.object(
            call.subprocess,
            "Popen",
            return_value=process,
        ) as popen, patch.object(
            call.uuid,
            "uuid4",
            return_value=SimpleNamespace(hex="launch-a"),
            create=True,
        ), patch.object(
            call,
            "utc_timestamp",
            return_value="2026-08-27T01:02:03.004Z",
        ), patch.object(
            call,
            "_listener_identity",
            side_effect=[
                {"pid": None, "createdAt": None},
                {"pid": None, "createdAt": None},
                {"pid": None, "createdAt": None},
                {"pid": None, "createdAt": None},
                {"pid": None, "createdAt": None},
                {"pid": None, "createdAt": None},
            ],
        ), patch.object(call.time, "sleep"):
            result = call._ensure_server()

        self.assertTrue(result)
        self.assertTrue(result.started)
        self.assertEqual("self_started", result.disposition)
        launch_environment = popen.call_args.kwargs["env"]
        self.assertEqual("launch-a", launch_environment["WPS_BRIDGE_LAUNCH_ID"])
        self.assertEqual(str(os.getpid()), launch_environment["WPS_BRIDGE_LAUNCHED_BY_PID"])
        self.assertEqual("456", str(result.health["serverPid"]))

    def test_spawn_response_with_incomplete_launch_identity_is_unknown(self):
        process = MagicMock()
        process.pid = 456
        process.poll.return_value = None
        incomplete_health = {
            "status": "ok",
            **current_service_identity(instance_id="new-instance"),
            "launchId": "launch-a",
            "pid": 456,
            "serverPid": 456,
        }
        with patch.object(
            call,
            "_health",
            side_effect=[None, None, incomplete_health],
        ), patch.object(call.os.path, "exists", return_value=True), patch.object(
            call,
            "server_log_path",
            return_value=(None, None),
        ), patch.object(
            call.subprocess,
            "Popen",
            return_value=process,
        ), patch.object(
            call.uuid,
            "uuid4",
            return_value=SimpleNamespace(hex="launch-a"),
        ), patch.object(call.time, "sleep"):
            result = call._ensure_server()

        self.assertFalse(result)
        self.assertEqual("unknown_owner", result.disposition)
        process.terminate.assert_called_once_with()

    def test_spawn_response_cannot_reuse_mismatched_launch_on_child_pid(self):
        process = MagicMock()
        process.pid = 456
        process.poll.return_value = None
        mismatched_health = {
            "status": "ok",
            **current_service_identity(instance_id="unexpected-instance"),
            "launchId": "other-launch",
            "pid": 456,
            "serverPid": 456,
            "launchedByPid": 999,
            "serverPath": call._server_path(),
            "launchStartedAt": "2026-08-27T00:00:00.000Z",
        }
        with patch.object(
            call,
            "_health",
            side_effect=[None, None, mismatched_health],
        ), patch.object(call.os.path, "exists", return_value=True), patch.object(
            call,
            "server_log_path",
            return_value=(None, None),
        ), patch.object(
            call.subprocess,
            "Popen",
            return_value=process,
        ), patch.object(
            call.uuid,
            "uuid4",
            return_value=SimpleNamespace(hex="launch-a"),
        ), patch.object(call.time, "sleep"):
            result = call._ensure_server()

        self.assertFalse(result)
        self.assertEqual("unknown_owner", result.disposition)
        process.terminate.assert_called_once_with()

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
            "launchId": "windows-launch",
            "pid": 789,
            "serverPid": 789,
            "launchedByPid": os.getpid(),
            "serverPath": call._server_path(),
            "launchStartedAt": "2026-08-27T01:02:03.004Z",
        }
        process = MagicMock()
        process.pid = 789
        process.poll.return_value = None

        class _StartupInfo:
            def __init__(self):
                self.dwFlags = 0

        with patch.object(call, "_health", side_effect=[None, None, ready_health]), patch.object(
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
        ) as popen, patch.object(
            call.uuid,
            "uuid4",
            return_value=SimpleNamespace(hex="windows-launch"),
        ), patch.object(
            call,
            "utc_timestamp",
            return_value="2026-08-27T01:02:03.004Z",
        ), patch.object(
            call,
            "_listener_identity",
            side_effect=[
                {"pid": None, "createdAt": None},
                {"pid": None, "createdAt": None},
                {"pid": None, "createdAt": None},
                {"pid": None, "createdAt": None},
                {"pid": 789, "createdAt": "child-time"},
                {"pid": 789, "createdAt": "child-time"},
            ],
        ), patch.object(call.time, "sleep"):
            result = call._ensure_server()

        self.assertTrue(result)
        self.assertTrue(result.started)
        self.assertEqual(expected["projectId"], result.health["projectId"])
        kwargs = popen.call_args.kwargs
        self.assertEqual(0x200, kwargs["creationflags"])
        self.assertEqual(0x01, kwargs["startupinfo"].dwFlags)
        self.assertNotIn("start_new_session", kwargs)

    def test_windows_startup_lock_retries_until_winner_releases_it(self):
        locking = MagicMock(
            side_effect=[OSError(errno.EACCES, "locked"), None, None]
        )
        fake_msvcrt = SimpleNamespace(
            LK_NBLCK=1,
            LK_UNLCK=2,
            locking=locking,
        )

        with patch.object(call.sys, "platform", "win32"), patch.dict(
            sys.modules,
            {"msvcrt": fake_msvcrt},
        ), patch.object(call.time, "sleep") as sleep:
            with call._startup_lock():
                pass

        self.assertEqual(1, locking.call_args_list[0].args[1])
        sleep.assert_called_once_with(0.05)

    def test_startup_lock_failure_is_reported_as_unknown_owner(self):
        with patch.object(call, "_health", return_value=None), patch.object(
            call,
            "_startup_lock",
            side_effect=OSError(errno.EACCES, "lock denied"),
        ), patch.object(call.subprocess, "Popen") as popen:
            result = call._ensure_server()

        self.assertFalse(result)
        self.assertEqual("BRIDGE_START_LOCK_FAILED", result.code)
        self.assertEqual("unknown_owner", result.disposition)
        popen.assert_not_called()

    def test_call_output_exposes_bridge_ensure_disposition(self):
        service = ServiceStartResult(
            True,
            health={"instanceId": "instance-a", "launchId": "launch-a", "serverPid": 321},
            disposition="reused",
        )
        with patch.object(call, "_load_params_from_args", return_value=("ping", {}, None)), patch.object(
            call,
            "_ensure_server",
            return_value=service,
        ), patch.object(
            call,
            "_post",
            return_value={"success": True},
        ), patch("builtins.print") as output, self.assertRaises(SystemExit) as exit_result:
            call.main()

        self.assertEqual(0, exit_result.exception.code)
        payload = json.loads(output.call_args.args[0])
        self.assertEqual("reused", payload["bridgeEnsure"]["disposition"])
        self.assertEqual("instance-a", payload["bridgeEnsure"]["instanceId"])


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
