import io
import json
import socket
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import server
from service_lifecycle import BridgeLifecycle, build_service_identity


class _Clock:
    def __init__(self):
        self.value = 0.0

    def __call__(self):
        return self.value

    def advance(self, seconds):
        self.value += seconds


class _RecordingController:
    def __init__(self, clock=None, duration=0):
        self.clock = clock
        self.duration = duration
        self.closed = False

    def execute(self, action, params, trace=None):
        if self.clock:
            self.clock.advance(self.duration)
        return {"success": True, "data": {"action": action}}

    def close(self):
        self.closed = True


def _handler(path, lifecycle, payload=None, headers=None):
    handler = object.__new__(server.Handler)
    handler.path = path
    raw = json.dumps(payload or {}).encode("utf-8")
    identity_headers = {
        "X-WPS-Bridge-Project-Id": lifecycle.identity.get("projectId", ""),
        "X-WPS-Bridge-Instance-Id": lifecycle.identity.get("instanceId", ""),
    }
    handler.headers = {
        "Content-Length": str(len(raw)),
        **identity_headers,
        **(headers or {}),
    }
    handler.rfile = io.BytesIO(raw)
    handler.server = SimpleNamespace(lifecycle=lifecycle)
    captured = {}
    handler._send = lambda obj, code=200: captured.update(result=obj, code=code)
    return handler, captured


class ShutdownEndpointTests(unittest.TestCase):
    def setUp(self):
        self.identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        self.lifecycle = BridgeLifecycle(self.identity, idle_timeout_seconds=60)

    def test_health_exposes_instance_identity(self):
        handler, captured = _handler("/health", self.lifecycle)

        handler.do_GET()

        self.assertEqual(200, captured["code"])
        self.assertEqual(self.identity["instanceId"], captured["result"]["instanceId"])
        self.assertEqual(self.identity["projectRoot"], captured["result"]["projectRoot"])
        self.assertIn("pid", captured["result"])

    def test_shutdown_requests_graceful_exit(self):
        handler, captured = _handler(
            "/shutdown",
            self.lifecycle,
            headers={"X-WPS-Bridge-Instance-Id": self.identity["instanceId"]},
        )

        handler.do_POST()

        self.assertEqual(202, captured["code"])
        self.assertTrue(self.lifecycle.should_stop())
        self.assertEqual("api", self.lifecycle.stop_reason)

    def test_shutdown_rejects_wrong_instance(self):
        handler, captured = _handler(
            "/shutdown",
            self.lifecycle,
            headers={"X-WPS-Bridge-Instance-Id": "wrong"},
        )

        handler.do_POST()

        self.assertEqual(409, captured["code"])
        self.assertFalse(self.lifecycle.should_stop())


class ServerLoopTests(unittest.TestCase):
    def test_windows_bind_uses_exclusive_address(self):
        httpd = object.__new__(server.BridgeHTTPServer)
        httpd.socket = Mock()
        httpd.allow_reuse_address = True
        with patch.object(server.os, "name", "nt"), patch.object(
            server.socket,
            "SO_EXCLUSIVEADDRUSE",
            4242,
            create=True,
        ), patch.object(server.HTTPServer, "server_bind") as parent_bind:
            server.BridgeHTTPServer.server_bind(httpd)

        self.assertFalse(httpd.allow_reuse_address)
        httpd.socket.setsockopt.assert_called_once_with(socket.SOL_SOCKET, 4242, 1)
        parent_bind.assert_called_once_with()

    def test_accepted_connection_gets_bounded_io_timeout(self):
        connection = Mock()
        httpd = object.__new__(server.BridgeHTTPServer)
        with patch.object(
            server.HTTPServer,
            "get_request",
            return_value=(connection, ("127.0.0.1", 12345)),
        ):
            result = httpd.get_request()

        self.assertEqual((connection, ("127.0.0.1", 12345)), result)
        connection.settimeout.assert_called_once_with(server.REQUEST_IO_TIMEOUT_SECONDS)

    def test_idle_loop_stops_without_real_sleep(self):
        clock = _Clock()
        lifecycle = BridgeLifecycle(
            build_service_identity("/workspace/current", "code-a"),
            idle_timeout_seconds=5,
            clock=clock,
        )

        class _Httpd:
            timeout = None
            calls = 0

            def handle_request(self):
                self.calls += 1
                clock.advance(5)

        httpd = _Httpd()

        reason = server.serve_until_stopped(httpd, lifecycle, poll_interval=0.01)

        self.assertEqual("idle_timeout", reason)
        self.assertEqual(1, httpd.calls)

    def test_long_action_refreshes_idle_only_after_it_finishes(self):
        clock = _Clock()
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=10, clock=clock)
        handler, captured = _handler(
            "/dispatch",
            lifecycle,
            payload={"action": "setCellValue", "params": {"cell": "A1", "value": 1}},
        )
        controller = _RecordingController(clock=clock, duration=20)

        with patch.object(server, "get_app_controller", return_value=controller):
            handler.do_POST()

        self.assertTrue(captured["result"]["success"])
        self.assertFalse(lifecycle.should_stop())

    def test_dispatch_rejects_missing_project_identity_before_controller(self):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=10)
        handler, captured = _handler(
            "/dispatch",
            lifecycle,
            payload={"action": "setCellValue", "params": {"cell": "A1", "value": 1}},
            headers={"X-WPS-Bridge-Project-Id": ""},
        )

        with patch.object(server, "get_app_controller") as controller:
            handler.do_POST()

        self.assertEqual(409, captured["code"])
        self.assertEqual("BRIDGE_PROJECT_MISMATCH", captured["result"]["code"])
        controller.assert_not_called()

    def test_dispatch_rejects_negative_content_length_without_reading_body(self):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=10)
        handler, captured = _handler(
            "/dispatch",
            lifecycle,
            payload={"action": "setCellValue"},
        )
        handler.headers["Content-Length"] = "-1"
        handler.rfile = Mock()

        handler.do_POST()

        self.assertEqual(400, captured["code"])
        self.assertEqual("INVALID_CONTENT_LENGTH", captured["result"]["code"])
        handler.rfile.read.assert_not_called()

    def test_dispatch_rejects_oversized_body_without_reading_it(self):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=10)
        handler, captured = _handler(
            "/dispatch",
            lifecycle,
            payload={"action": "setCellValue"},
        )
        handler.headers["Content-Length"] = str(server.MAX_REQUEST_BYTES + 1)
        handler.rfile = Mock()

        handler.do_POST()

        self.assertEqual(413, captured["code"])
        self.assertEqual("REQUEST_TOO_LARGE", captured["result"]["code"])
        handler.rfile.read.assert_not_called()

    def test_controller_cleanup_continues_after_one_close_failure(self):
        first = _RecordingController()
        second = _RecordingController()
        first.close = lambda: (_ for _ in ()).throw(RuntimeError("close failed"))

        server._cleanup_controllers({"first": first, "second": second})

        self.assertTrue(second.closed)


class ServerMainCleanupTests(unittest.TestCase):
    def _run_main(self, serve_side_effect=None):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        httpd = Mock()
        if serve_side_effect is None:
            serve = patch.object(server, "serve_until_stopped", return_value="idle_timeout")
        else:
            serve = patch.object(server, "serve_until_stopped", side_effect=serve_side_effect)
        patches = (
            patch.object(server, "new_service_identity", return_value=identity),
            patch.object(server, "configured_idle_timeout", return_value=60),
            patch.object(server, "BridgeHTTPServer", return_value=httpd),
            patch.object(server, "_install_signal_handlers", return_value={}),
            patch.object(server, "_restore_signal_handlers"),
            patch.object(server, "_cleanup_controllers"),
            patch("builtins.print"),
            serve,
        )
        entered = []
        try:
            for context in patches:
                entered.append(context.__enter__())
            server.main()
        finally:
            for context in reversed(patches):
                context.__exit__(None, None, None)
        return httpd, entered[4], entered[5]

    def test_normal_exit_always_closes_socket_and_controllers(self):
        httpd, restore, cleanup = self._run_main()

        httpd.server_close.assert_called_once_with()
        cleanup.assert_called_once_with()
        restore.assert_called_once_with({})

    def test_unexpected_loop_error_still_closes_socket_and_controllers(self):
        with self.assertRaisesRegex(RuntimeError, "loop failed"):
            self._run_main(serve_side_effect=RuntimeError("loop failed"))


if __name__ == "__main__":
    unittest.main()
