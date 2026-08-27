import io
import json
import socket
from socketserver import BaseRequestHandler
import threading
import time
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch
from urllib.request import Request, urlopen

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


def _post_dispatch(httpd, lifecycle, action, trace_id):
    body = json.dumps({"action": action, "params": {}}).encode("utf-8")
    request = Request(
        f"http://127.0.0.1:{httpd.server_port}/dispatch",
        data=body,
        headers={
            "Content-Type": "application/json",
            "X-WPS-Bridge-Project-Id": lifecycle.identity["projectId"],
            "X-WPS-Bridge-Instance-Id": lifecycle.identity["instanceId"],
            "X-WPS-Trace-Id": trace_id,
        },
        method="POST",
    )
    with urlopen(request, timeout=2) as response:
        return json.load(response)


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

    def test_shutdown_starts_before_acceptance_response_is_written(self):
        shutdown, _ = _handler("/shutdown", self.lifecycle)
        action, action_result = _handler(
            "/dispatch",
            self.lifecycle,
            payload={"action": "ping", "params": {}},
        )
        response_started = threading.Event()
        release_response = threading.Event()

        def blocking_send(_result, _code=200):
            response_started.set()
            release_response.wait(timeout=1)

        shutdown._send = blocking_send
        shutdown_thread = threading.Thread(target=shutdown.do_POST)
        try:
            with patch.object(server, "dispatch") as fake_dispatch:
                shutdown_thread.start()
                self.assertTrue(response_started.wait(timeout=1))

                action.do_POST()

                self.assertEqual(503, action_result["code"])
                self.assertEqual(
                    "BRIDGE_SHUTTING_DOWN",
                    action_result["result"]["code"],
                )
                fake_dispatch.assert_not_called()
        finally:
            release_response.set()
            shutdown_thread.join(timeout=2)


class ServerLoopTests(unittest.TestCase):
    def test_handler_concurrency_is_bounded(self):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=0)
        count_lock = threading.Lock()
        two_entered = threading.Event()
        release = threading.Event()
        active = 0
        maximum = 0

        class BlockingHandler(BaseRequestHandler):
            def handle(self):
                nonlocal active, maximum
                with count_lock:
                    active += 1
                    maximum = max(maximum, active)
                    if active == 2:
                        two_entered.set()
                try:
                    release.wait(timeout=1)
                finally:
                    with count_lock:
                        active -= 1

        httpd = server.BridgeHTTPServer(
            ("127.0.0.1", 0),
            BlockingHandler,
            lifecycle,
            max_handler_threads=2,
        )
        serving = threading.Thread(
            target=server.serve_until_stopped,
            args=(httpd, lifecycle),
        )
        serving.start()
        clients = []
        try:
            for _ in range(6):
                client = socket.create_connection(httpd.server_address, timeout=1)
                clients.append(client)
            self.assertTrue(two_entered.wait(timeout=1))
            time.sleep(0.05)
            with count_lock:
                observed_maximum = maximum
        finally:
            release.set()
            for client in clients:
                client.close()
            lifecycle.request_stop("test_cleanup")
            serving.join(timeout=2)
            httpd.server_close()

        self.assertEqual(2, observed_maximum)

    def test_server_close_waits_for_active_handler(self):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=0)
        entered = threading.Event()
        release = threading.Event()

        class BlockingHandler(BaseRequestHandler):
            def handle(self):
                entered.set()
                release.wait(timeout=1)

        httpd = server.BridgeHTTPServer(
            ("127.0.0.1", 0),
            BlockingHandler,
            lifecycle,
        )
        serving = threading.Thread(
            target=server.serve_until_stopped,
            args=(httpd, lifecycle),
        )
        serving.start()
        client = socket.create_connection(httpd.server_address, timeout=1)
        closing = threading.Thread(target=httpd.server_close)
        try:
            self.assertTrue(entered.wait(timeout=1))
            lifecycle.request_stop("test_cleanup")
            serving.join(timeout=2)
            closing.start()
            time.sleep(0.05)

            self.assertTrue(closing.is_alive())

            release.set()
            closing.join(timeout=2)
        finally:
            release.set()
            client.close()
            if closing.ident is not None:
                closing.join(timeout=2)
            else:
                httpd.server_close()

        self.assertFalse(closing.is_alive())

    def test_incomplete_connection_does_not_block_health(self):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=0)
        httpd = server.BridgeHTTPServer(("127.0.0.1", 0), server.Handler, lifecycle)
        serving = threading.Thread(
            target=server.serve_until_stopped,
            args=(httpd, lifecycle),
        )
        serving.start()
        incomplete = socket.create_connection(httpd.server_address, timeout=1)

        try:
            incomplete.sendall(b"GET /health HTTP/1.1\r\nHost: localhost\r\n")
            time.sleep(0.05)

            with urlopen(
                f"http://127.0.0.1:{httpd.server_port}/health",
                timeout=0.3,
            ) as response:
                health = json.load(response)

            self.assertEqual("ok", health["status"])
        finally:
            incomplete.close()
            lifecycle.request_stop("test_cleanup")
            serving.join(timeout=2)
            httpd.server_close()

        self.assertFalse(serving.is_alive())

    def test_concurrent_dispatches_enter_action_execution_one_at_a_time(self):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=0)
        httpd = server.BridgeHTTPServer(("127.0.0.1", 0), server.Handler, lifecycle)
        serving = threading.Thread(
            target=server.serve_until_stopped,
            args=(httpd, lifecycle),
        )
        serving.start()
        entered = threading.Event()
        release = threading.Event()
        count_lock = threading.Lock()
        active = 0
        maximum = 0
        responses = []

        def blocking_dispatch(action, params, app=None, trace=None, prepared=None):
            nonlocal active, maximum
            with count_lock:
                active += 1
                maximum = max(maximum, active)
                entered.set()
            try:
                release.wait(timeout=1)
                return {"success": True, "data": action}
            finally:
                with count_lock:
                    active -= 1

        def call(action, trace_id):
            responses.append(_post_dispatch(httpd, lifecycle, action, trace_id))

        first_trace = "act-20260827-44444444444444444444444444444444"
        second_trace = "act-20260827-55555555555555555555555555555555"
        first = threading.Thread(target=call, args=("ping", first_trace))
        second = threading.Thread(target=call, args=("wireCheck", second_trace))
        try:
            with patch.object(server, "dispatch", side_effect=blocking_dispatch):
                first.start()
                self.assertTrue(entered.wait(timeout=1))
                second.start()
                time.sleep(0.1)
                with count_lock:
                    observed_maximum = maximum
                release.set()
                first.join(timeout=2)
                second.join(timeout=2)
        finally:
            release.set()
            lifecycle.request_stop("test_cleanup")
            serving.join(timeout=2)
            httpd.server_close()

        self.assertEqual(1, observed_maximum)
        self.assertEqual(2, len(responses))
        response_traces = {result["data"]: result["traceId"] for result in responses}
        self.assertEqual(first_trace, response_traces["ping"])
        self.assertEqual(second_trace, response_traces["wireCheck"])

    def test_blocked_dispatch_does_not_block_health(self):
        trace_id = "act-20260827-11111111111111111111111111111111"
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=0)
        httpd = server.BridgeHTTPServer(("127.0.0.1", 0), server.Handler, lifecycle)
        serving = threading.Thread(
            target=server.serve_until_stopped,
            args=(httpd, lifecycle),
        )
        serving.start()
        entered = threading.Event()
        release = threading.Event()
        responses = []

        def blocking_dispatch(action, params, app=None, trace=None, prepared=None):
            entered.set()
            release.wait(timeout=1)
            return {"success": True, "data": action}

        action_call = threading.Thread(
            target=lambda: responses.append(
                _post_dispatch(httpd, lifecycle, "ping", trace_id)
            )
        )
        try:
            with patch.object(server, "dispatch", side_effect=blocking_dispatch):
                action_call.start()
                self.assertTrue(entered.wait(timeout=1))

                started = time.monotonic()
                with urlopen(
                    f"http://127.0.0.1:{httpd.server_port}/health",
                    timeout=0.3,
                ) as response:
                    health = json.load(response)
                elapsed = time.monotonic() - started

                release.set()
                action_call.join(timeout=2)
        finally:
            release.set()
            lifecycle.request_stop("test_cleanup")
            serving.join(timeout=2)
            httpd.server_close()

        self.assertLess(elapsed, 0.3)
        self.assertEqual("running", health["state"])
        self.assertEqual(trace_id, health["activeAction"]["traceId"])
        self.assertEqual("ping", health["activeAction"]["action"])
        self.assertEqual(1, len(responses))

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

    def test_invalid_action_does_not_refresh_idle_deadline(self):
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
            payload={"action": "notARealAction", "params": {}},
        )
        clock.advance(9)

        handler.do_POST()
        clock.advance(1)

        self.assertEqual("UNKNOWN_ACTION", captured["result"]["code"])
        self.assertTrue(lifecycle.should_stop())

    def test_shutdown_allows_active_action_and_rejects_queued_action(self):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=0)
        first, first_result = _handler(
            "/dispatch",
            lifecycle,
            payload={"action": "ping", "params": {}},
            headers={
                "X-WPS-Trace-Id": "act-20260827-22222222222222222222222222222222"
            },
        )
        second, second_result = _handler(
            "/dispatch",
            lifecycle,
            payload={"action": "ping", "params": {}},
            headers={
                "X-WPS-Trace-Id": "act-20260827-33333333333333333333333333333333"
            },
        )
        shutdown, shutdown_result = _handler("/shutdown", lifecycle)
        second_raw = second.rfile.getvalue()
        second_body_read = threading.Event()
        entered = threading.Event()
        release = threading.Event()

        def read_second_body(length):
            second_body_read.set()
            return second_raw[:length]

        second.rfile = SimpleNamespace(read=read_second_body)

        def blocking_dispatch(action, params, app=None, trace=None, prepared=None):
            entered.set()
            release.wait(timeout=1)
            return {"success": True, "data": action}

        first_thread = threading.Thread(target=first.do_POST)
        second_thread = threading.Thread(target=second.do_POST)
        try:
            with patch.object(
                server,
                "dispatch",
                side_effect=blocking_dispatch,
            ) as fake_dispatch:
                first_thread.start()
                self.assertTrue(entered.wait(timeout=1))
                second_thread.start()
                self.assertTrue(second_body_read.wait(timeout=1))

                shutdown.do_POST()
                release.set()
                first_thread.join(timeout=2)
                second_thread.join(timeout=2)

                self.assertEqual(1, fake_dispatch.call_count)
        finally:
            release.set()

        self.assertEqual(202, shutdown_result["code"])
        self.assertTrue(first_result["result"]["success"])
        self.assertEqual(503, second_result["code"])
        self.assertEqual(
            "BRIDGE_SHUTTING_DOWN",
            second_result["result"]["code"],
        )

    def test_slow_request_body_does_not_hold_action_execution_lock(self):
        identity = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-a",
        )
        lifecycle = BridgeLifecycle(identity, idle_timeout_seconds=0)
        slow, slow_result = _handler(
            "/dispatch",
            lifecycle,
            payload={"action": "ping", "params": {}},
        )
        ready, ready_result = _handler(
            "/dispatch",
            lifecycle,
            payload={"action": "wireCheck", "params": {}},
        )
        slow_raw = slow.rfile.getvalue()
        read_started = threading.Event()
        release_body = threading.Event()

        def read_slow_body(length):
            read_started.set()
            release_body.wait(timeout=1)
            return slow_raw[:length]

        slow.rfile = SimpleNamespace(read=read_slow_body)
        slow_thread = threading.Thread(target=slow.do_POST)
        try:
            with patch.object(
                server,
                "dispatch",
                return_value={"success": True},
            ) as fake_dispatch:
                slow_thread.start()
                self.assertTrue(read_started.wait(timeout=1))

                ready.do_POST()

                self.assertTrue(ready_result["result"]["success"])
                self.assertEqual(1, fake_dispatch.call_count)
                release_body.set()
                slow_thread.join(timeout=2)
                self.assertEqual(2, fake_dispatch.call_count)
        finally:
            release_body.set()

        self.assertTrue(slow_result["result"]["success"])

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
