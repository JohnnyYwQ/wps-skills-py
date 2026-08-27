import sys
import errno
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))
import service

from service_lifecycle import (
    ServiceStartResult,
    build_service_identity,
    current_service_identity,
)


class ServiceCliTests(unittest.TestCase):
    def test_start_exposes_bridge_discovery_disposition(self):
        ready = ServiceStartResult(
            True,
            health={"instanceId": "instance-a"},
            disposition="reused",
        )
        with patch.object(service.call, "_ensure_server", return_value=ready):
            result = service.start()

        self.assertEqual("reused", result["disposition"])

    def test_windows_wait_failures_are_conservatively_treated_as_running(self):
        self.assertFalse(service._windows_wait_result_is_running(0x00000000))
        self.assertTrue(service._windows_wait_result_is_running(0x00000102))
        self.assertTrue(service._windows_wait_result_is_running(0xFFFFFFFF))

    def test_windows_only_invalid_pid_error_means_not_running(self):
        self.assertFalse(service._windows_open_error_means_running(87))
        self.assertTrue(service._windows_open_error_means_running(5))
        self.assertTrue(service._windows_open_error_means_running(8))

    def test_posix_transient_pid_probe_error_is_conservatively_running(self):
        with patch.object(
            service.os,
            "kill",
            side_effect=OSError(errno.EINTR, "interrupted"),
        ):
            self.assertTrue(service._process_is_running(4321))

    def test_status_does_not_start_a_stopped_service(self):
        with patch.object(service.call, "_health", return_value=None), patch.object(
            service.call,
            "_ensure_server",
        ) as ensure:
            result = service.status()

        self.assertTrue(result["success"])
        self.assertEqual("stopped", result["state"])
        ensure.assert_not_called()

    def test_status_reports_unresponsive_service_instead_of_stopped(self):
        probe = service.call.HealthProbe("unresponsive", error="timed out")
        with patch.object(service.call, "_health", return_value=probe):
            result = service.status()

        self.assertTrue(result["success"])
        self.assertEqual("unresponsive", result["state"])
        self.assertFalse(result["reusable"])

    def test_stop_does_not_claim_unresponsive_service_is_stopped(self):
        probe = service.call.HealthProbe("unresponsive", error="timed out")
        with patch.object(service.call, "_health", return_value=probe), patch.object(
            service,
            "_post_shutdown",
        ) as shutdown:
            result = service.stop()

        self.assertFalse(result["success"])
        self.assertEqual("unresponsive", result["state"])
        shutdown.assert_not_called()

    def test_json_null_health_is_unhealthy_not_stopped(self):
        probe = service.call.HealthProbe("responded", health=None)
        with patch.object(service.call, "_health", return_value=probe):
            status_result = service.status()
            stop_result = service.stop()

        self.assertEqual("unhealthy", status_result["state"])
        self.assertFalse(stop_result["success"])
        self.assertEqual("unhealthy", stop_result["state"])

    def test_stop_is_idempotent_when_service_is_absent(self):
        with patch.object(service.call, "_health", return_value=None):
            result = service.stop()

        self.assertTrue(result["success"])
        self.assertEqual("stopped", result["state"])

    def test_stop_refuses_a_foreign_checkout_without_takeover(self):
        foreign = {
            "status": "ok",
            **build_service_identity(
                "/workspace/other",
                "code-a",
                instance_id="foreign-instance",
            ),
        }
        with patch.object(service.call, "_health", return_value=foreign), patch.object(
            service,
            "_post_shutdown",
        ) as shutdown:
            result = service.stop()

        self.assertFalse(result["success"])
        self.assertEqual("foreign", result["state"])
        shutdown.assert_not_called()

    def test_partial_legacy_identity_is_refused_without_key_error(self):
        expected = current_service_identity()
        partial = {
            "status": "ok",
            "projectId": expected["projectId"],
            "projectRoot": expected["projectRoot"],
        }
        with patch.object(service.call, "_health", return_value=partial), patch.object(
            service,
            "_post_shutdown",
        ) as shutdown:
            result = service.stop()

        self.assertFalse(result["success"])
        self.assertEqual("legacy", result["state"])
        shutdown.assert_not_called()

    def test_takeover_cannot_bypass_partial_legacy_refusal(self):
        expected = current_service_identity()
        partial = {
            "status": "ok",
            "service": "wps-skills-bridge",
            "projectId": "foreign-project",
            "instanceId": "foreign-instance",
            "codeFingerprint": expected["codeFingerprint"],
        }
        with patch.object(service.call, "_health", return_value=partial), patch.object(
            service,
            "_post_shutdown",
        ) as shutdown:
            result = service.stop(takeover=True)

        self.assertFalse(result["success"])
        self.assertEqual("legacy", result["state"])
        shutdown.assert_not_called()

    def test_takeover_can_cooperatively_stop_complete_foreign_bridge(self):
        foreign = {
            "status": "ok",
            **build_service_identity(
                "/workspace/other",
                "code-a",
                instance_id="foreign-instance",
            ),
            "pid": 4321,
        }
        with patch.object(service.call, "_health", return_value=foreign), patch.object(
            service,
            "_post_shutdown",
            return_value={"success": True, "status": "stopping"},
        ) as shutdown, patch.object(
            service,
            "_wait_until_instance_stops",
            return_value=True,
        ) as wait_for_stop:
            result = service.stop(takeover=True)

        self.assertTrue(result["success"])
        shutdown.assert_called_once_with(foreign)
        wait_for_stop.assert_called_once_with("foreign-instance", pid=4321)

    def test_same_checkout_stale_service_can_be_stopped(self):
        expected = current_service_identity()
        stale = {
            "status": "ok",
            **build_service_identity(
                expected["projectRoot"],
                "old-code",
                instance_id="stale-instance",
            ),
        }
        with patch.object(service.call, "_health", side_effect=[stale, None]), patch.object(
            service,
            "_post_shutdown",
            return_value={"success": True, "status": "stopping"},
        ) as shutdown, patch.object(service.time, "sleep"):
            result = service.stop()

        self.assertTrue(result["success"])
        self.assertEqual("stopped", result["state"])
        shutdown.assert_called_once_with(stale)

    def test_shutdown_connection_race_is_success_when_original_pid_exits(self):
        health = {
            "status": "ok",
            **current_service_identity(instance_id="racing-instance"),
            "pid": 4321,
        }
        with patch.object(service.call, "_health", return_value=health), patch.object(
            service,
            "_post_shutdown",
            return_value={"success": False, "error": "connection refused"},
        ), patch.object(
            service,
            "_wait_until_instance_stops",
            return_value=True,
        ) as wait_for_stop:
            result = service.stop()

        self.assertTrue(result["success"])
        self.assertTrue(result["shutdownRace"])
        wait_for_stop.assert_called_once_with("racing-instance", pid=4321)

    def test_already_stopping_instance_is_waited_without_second_shutdown(self):
        health = {
            "status": "ok",
            **current_service_identity(instance_id="stopping-instance"),
            "pid": 4321,
            "shutdownRequested": True,
        }
        with patch.object(service.call, "_health", return_value=health), patch.object(
            service,
            "_wait_until_instance_stops",
            return_value=True,
        ) as wait_for_stop, patch.object(service, "_post_shutdown") as shutdown:
            result = service.stop()

        self.assertTrue(result["success"])
        self.assertTrue(result["alreadyStopping"])
        shutdown.assert_not_called()
        wait_for_stop.assert_called_once_with("stopping-instance", pid=4321)

    def test_wait_does_not_trust_missing_health_while_original_pid_is_alive(self):
        with patch.object(service.call, "_health", return_value=None), patch.object(
            service,
            "_process_is_running",
            side_effect=[True, False],
        ) as process_running, patch.object(service.time, "sleep"):
            stopped = service._wait_until_instance_stops(
                "instance-a",
                pid=4321,
                timeout=1,
            )

        self.assertTrue(stopped)
        self.assertEqual(2, process_running.call_count)

    def test_restart_never_starts_when_stop_fails(self):
        with patch.object(
            service,
            "stop",
            return_value={"success": False, "code": "BRIDGE_STOP_FAILED"},
        ), patch.object(service, "start") as start:
            result = service.restart()

        self.assertFalse(result["success"])
        start.assert_not_called()

    def test_restart_reports_previous_and_new_instances(self):
        started = {
            "success": True,
            "state": "running",
            "service": {"instanceId": "new-instance"},
        }
        with patch.object(
            service,
            "stop",
            return_value={
                "success": True,
                "state": "stopped",
                "stoppedInstanceId": "old-instance",
            },
        ), patch.object(service, "start", return_value=started):
            result = service.restart()

        self.assertTrue(result["success"])
        self.assertTrue(result["restarted"])
        self.assertEqual("old-instance", result["previousInstanceId"])
        self.assertEqual("new-instance", result["service"]["instanceId"])


if __name__ == "__main__":
    unittest.main()
