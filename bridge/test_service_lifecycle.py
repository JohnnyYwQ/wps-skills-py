import tempfile
import subprocess
import unittest
import os
from pathlib import Path
from unittest.mock import Mock, patch

from service_lifecycle import (
    BridgeLifecycle,
    build_service_identity,
    code_fingerprint,
    inspect_health,
    new_service_identity,
    stop_line_process,
)


class _Clock:
    def __init__(self):
        self.value = 0.0

    def __call__(self):
        return self.value

    def advance(self, seconds):
        self.value += seconds


class ServiceIdentityTests(unittest.TestCase):
    def setUp(self):
        self.expected = build_service_identity("/workspace/current", "code-a")

    def test_only_exact_service_identity_is_reusable(self):
        health = {
            "status": "ok",
            **build_service_identity("/workspace/current", "code-a", instance_id="instance-a"),
        }

        inspection = inspect_health(health, expected=self.expected)

        self.assertEqual("current", inspection.state)
        self.assertTrue(inspection.reusable)

    def test_legacy_health_without_identity_is_rejected(self):
        inspection = inspect_health({"status": "ok"}, expected=self.expected)

        self.assertEqual("legacy", inspection.state)
        self.assertFalse(inspection.reusable)

    def test_other_checkout_is_rejected(self):
        other = build_service_identity("/workspace/other", "code-a", instance_id="instance-b")

        inspection = inspect_health({"status": "ok", **other}, expected=self.expected)

        self.assertEqual("foreign", inspection.state)
        self.assertFalse(inspection.reusable)

    def test_project_root_must_match_even_if_project_id_is_copied(self):
        copied = build_service_identity(
            "/workspace/current",
            "code-a",
            instance_id="instance-b",
        )
        copied["projectRoot"] = "/workspace/other"

        inspection = inspect_health({"status": "ok", **copied}, expected=self.expected)

        self.assertEqual("foreign", inspection.state)
        self.assertFalse(inspection.reusable)

    def test_stale_code_in_the_same_checkout_is_rejected(self):
        stale = build_service_identity("/workspace/current", "code-old", instance_id="instance-old")

        inspection = inspect_health({"status": "ok", **stale}, expected=self.expected)

        self.assertEqual("stale", inspection.state)
        self.assertFalse(inspection.reusable)
        self.assertTrue(inspection.same_project)

    def test_shutting_down_instance_is_not_reused(self):
        health = {
            "status": "ok",
            **build_service_identity("/workspace/current", "code-a", instance_id="instance-a"),
            "shutdownRequested": True,
        }

        inspection = inspect_health(health, expected=self.expected)

        self.assertEqual("stopping", inspection.state)
        self.assertFalse(inspection.reusable)

    def test_runtime_fingerprint_changes_when_runtime_code_changes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "bridge").mkdir()
            (root / "scripts").mkdir()
            runtime_file = root / "bridge" / "server.py"
            runtime_file.write_text("version = 1\n", encoding="utf-8")
            (root / "scripts" / "call.py").write_text("call = 1\n", encoding="utf-8")
            before = code_fingerprint(root)

            runtime_file.write_text("version = 2\n", encoding="utf-8")
            after = code_fingerprint(root)

        self.assertNotEqual(before, after)

    def test_runtime_fingerprint_includes_linux_vendored_dependencies(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "bridge").mkdir()
            (root / "scripts").mkdir()
            package = root / "vendor" / "openpyxl"
            package.mkdir(parents=True)
            dependency = package / "__init__.py"
            dependency.write_text("version = 1\n", encoding="utf-8")
            before = code_fingerprint(root)

            dependency.write_text("version = 2\n", encoding="utf-8")
            after = code_fingerprint(root)

        self.assertNotEqual(before, after)

    def test_server_identity_carries_caller_launch_context(self):
        launch_context = {
            "WPS_BRIDGE_LAUNCH_ID": "launch-a",
            "WPS_BRIDGE_LAUNCHED_BY_PID": "321",
            "WPS_BRIDGE_SERVER_PATH": "/workspace/current/bridge/server.py",
            "WPS_BRIDGE_LAUNCH_STARTED_AT": "2026-08-27T01:02:03.004Z",
        }

        with patch.dict(os.environ, launch_context, clear=False):
            identity = new_service_identity()

        self.assertEqual("launch-a", identity["launchId"])
        self.assertEqual(321, identity["launchedByPid"])
        self.assertEqual(
            "/workspace/current/bridge/server.py",
            identity["serverPath"],
        )
        self.assertEqual(
            "2026-08-27T01:02:03.004Z",
            identity["launchStartedAt"],
        )


class BridgeLifecycleTests(unittest.TestCase):
    def setUp(self):
        self.clock = _Clock()
        identity = build_service_identity("/workspace/current", "code-a")
        self.lifecycle = BridgeLifecycle(
            identity=identity,
            idle_timeout_seconds=10,
            clock=self.clock,
        )

    def test_idle_timeout_requests_shutdown(self):
        self.clock.advance(10)

        self.assertTrue(self.lifecycle.should_stop())
        self.assertEqual("idle_timeout", self.lifecycle.stop_reason)

    def test_completed_action_resets_idle_deadline(self):
        self.clock.advance(9)
        self.lifecycle.mark_action_completed()
        self.clock.advance(9)

        self.assertFalse(self.lifecycle.should_stop())

    def test_active_action_can_exceed_idle_timeout_without_stopping(self):
        with self.lifecycle.action_execution("setCellValue", "trace-active") as admitted:
            self.assertTrue(admitted)
            self.clock.advance(20)

            self.assertFalse(self.lifecycle.should_stop())

        self.clock.advance(9)
        self.assertFalse(self.lifecycle.should_stop())

    def test_health_snapshot_does_not_reset_idle_deadline(self):
        self.clock.advance(9)
        snapshot = self.lifecycle.health_snapshot()
        self.clock.advance(1)

        self.assertEqual("ok", snapshot["status"])
        self.assertTrue(self.lifecycle.should_stop())

    def test_health_distinguishes_idle_running_and_stopping(self):
        self.assertEqual("idle", self.lifecycle.health_snapshot()["state"])

        with self.lifecycle.action_execution("setCellValue", "trace-active"):
            running = self.lifecycle.health_snapshot()

            self.assertEqual("running", running["state"])
            self.assertEqual("trace-active", running["activeAction"]["traceId"])
            self.assertEqual("setCellValue", running["activeAction"]["action"])
            self.assertIn("startedAt", running["activeAction"])
            self.assertEqual(
                {"traceId", "action", "startedAt"},
                set(running["activeAction"]),
            )

        self.lifecycle.request_stop("api")
        self.assertEqual("stopping", self.lifecycle.health_snapshot()["state"])

    def test_explicit_shutdown_is_idempotent(self):
        self.lifecycle.request_stop("api")
        self.lifecycle.request_stop("api")

        self.assertTrue(self.lifecycle.should_stop())
        self.assertEqual("api", self.lifecycle.stop_reason)

    def test_signal_stop_is_deferred_until_normal_lifecycle_poll(self):
        self.lifecycle.request_stop_from_signal("signal:SIGTERM")

        self.assertTrue(self.lifecycle.should_stop())
        self.assertEqual("signal:SIGTERM", self.lifecycle.stop_reason)

    def test_zero_idle_timeout_disables_automatic_shutdown(self):
        lifecycle = BridgeLifecycle(
            identity=build_service_identity("/workspace/current", "code-a"),
            idle_timeout_seconds=0,
            clock=self.clock,
        )
        self.clock.advance(100000)

        self.assertFalse(lifecycle.should_stop())


class LineProcessShutdownTests(unittest.TestCase):
    def test_timeout_escalates_from_exit_to_terminate_then_kill(self):
        process = Mock()
        process.poll.return_value = None
        process.wait.side_effect = [
            subprocess.TimeoutExpired("powershell", 2),
            subprocess.TimeoutExpired("powershell", 1),
            0,
        ]

        stop_line_process(process, graceful_timeout=2, forced_timeout=1)

        process.stdin.write.assert_called_once_with("EXIT\n")
        process.stdin.flush.assert_called_once_with()
        process.terminate.assert_called_once_with()
        process.kill.assert_called_once_with()
        self.assertEqual(3, process.wait.call_count)


if __name__ == "__main__":
    unittest.main()
