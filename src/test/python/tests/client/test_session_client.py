import os
from pathlib import Path
import sys
import tempfile
import time
import unittest

from wps_skills.client import session_client
from wps_skills.client.session_client import ActionFailed, SessionClient, SessionClientError


SOURCE = Path(session_client.__file__).resolve().parents[2]
FIXTURE = Path(__file__).resolve().parents[3] / "resources/wps_skills/client/session_host_fixture.py"


class SessionClientTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.events = Path(self.directory.name) / "events.txt"

    def client(self, mode="normal"):
        return SessionClient(
            [sys.executable, str(FIXTURE), mode, str(self.events)],
            application="word", cwd=self.directory.name,
            env=dict(os.environ, PYTHONPATH=str(SOURCE), PYTHONIOENCODING="utf-8"),
            timeout=5,
        )

    @staticmethod
    def address(action):
        return {"app": "word", "action": action}

    def test_real_host_round_trips_unicode_and_closes_idempotently(self):
        client = self.client()
        with client:
            response = client.call(self.address("writeContent"), {"text": "中文 😀"})
            self.assertEqual("中文 😀", response["data"]["echo"]["text"])
            self.assertTrue(client.can_execute)
        self.assertEqual("succeeded", client.session_outcome["outcome"])
        self.assertEqual("client_close", client.session_outcome["reason"])
        self.assertEqual(client.session_outcome, client.close())
        self.assertFalse(client.can_execute)

    def test_terminal_notification_preserves_real_action_error_and_stops_next_write(self):
        client = self.client()
        with self.assertRaises(ActionFailed) as caught:
            with client:
                client.call(self.address("terminal"), {})
                client.call(self.address("must_not_run"), {})
        self.assertEqual("unknown", caught.exception.response["outcome"])
        self.assertEqual("CONTENT_VERIFICATION_FAILED", caught.exception.response["error"]["code"])
        self.assertEqual("host_failure", client.session_outcome["reason"])
        self.assertEqual("terminal\n", self.events.read_text())
        with self.assertRaises(SessionClientError):
            client.call(self.address("must_not_run"), {})

    def test_definite_failure_requires_explicit_caller_decision_to_continue(self):
        with self.client() as client:
            with self.assertRaises(ActionFailed):
                client.call(self.address("fail"), {})
            self.assertTrue(client.can_execute)
            self.assertEqual("succeeded", client.call(self.address("inspectDocument"), {})["outcome"])

    def test_lost_action_response_is_never_replayed(self):
        client = self.client("eof")
        with self.assertRaises(SessionClientError) as caught:
            with client:
                client.call(self.address("writeContent"), {})
        self.assertTrue(caught.exception.may_have_effect)
        self.assertIsNone(client.last_response)
        self.assertFalse(client.can_execute)
        self.assertEqual("writeContent\n", self.events.read_text())

    def test_wrong_session_response_is_not_accepted(self):
        client = self.client("mismatch")
        with self.assertRaises(SessionClientError) as caught:
            with client:
                client.call(self.address("writeContent"), {})
        self.assertTrue(caught.exception.may_have_effect)
        self.assertIsNone(client.last_response)

    def test_timeout_disconnects_host_without_replaying_the_request(self):
        client = self.client("timeout").start()
        client._timeout = 0.1
        with self.assertRaises(SessionClientError) as caught:
            client.call(self.address("writeContent"), {})
        self.assertTrue(caught.exception.may_have_effect)
        self.assertEqual("writeContent\n", self.events.read_text())
        self.assertIsNotNone(client._process.poll())

    def test_interruption_disconnects_instead_of_sending_another_request(self):
        from unittest.mock import patch

        client = self.client("timeout").start()
        with patch.object(client, "_receive", side_effect=KeyboardInterrupt):
            with self.assertRaises(KeyboardInterrupt):
                client.call(self.address("writeContent"), {})
        self.assertFalse(client.can_execute)
        self.assertIsNotNone(client._process.poll())
        self.assertEqual("writeContent\n", self.events.read_text())

    def test_missing_cleanup_record_does_not_hide_action_error(self):
        client = self.client("missing_closed")
        with self.assertRaises(ActionFailed) as caught:
            with client:
                client.call(self.address("writeContent"), {})
        self.assertEqual("CONTENT_VERIFICATION_FAILED", caught.exception.response["error"]["code"])
        self.assertIsNotNone(client.cleanup_error)
        self.assertIsNone(client.session_outcome)

    def test_cleanup_failure_does_not_hide_action_error(self):
        client = self.client("cleanup_failed")
        with self.assertRaises(ActionFailed):
            with client:
                client.call(self.address("terminal"), {})
        self.assertEqual("failed", client.session_outcome["outcome"])
        self.assertIsNotNone(client.cleanup_error)

    def test_cleanup_failure_after_success_preserves_completed_response(self):
        client = self.client("cleanup_failed")
        with self.assertRaises(SessionClientError):
            with client:
                client.call(self.address("writeContent"), {})
        self.assertEqual("succeeded", client.last_response["outcome"])
        self.assertEqual("failed", client.session_outcome["outcome"])

    def test_stderr_is_drained_without_blocking_host(self):
        client = self.client("stderr")
        with client:
            client.call(self.address("inspectDocument"), {})
        self.assertIn("fixture diagnostic output", client.stderr)

    def test_startup_failure_has_no_possible_action_effect(self):
        client = self.client("startup")
        with self.assertRaises(SessionClientError) as caught:
            client.start()
        self.assertFalse(caught.exception.may_have_effect)
        self.assertIn("fixture startup failed", client.stderr)

    def test_idle_closure_is_consumed_before_another_request(self):
        client = self.client("idle").start()
        # Wait for the fixture Host to exit, not a guessed wall-clock delay.
        deadline = time.monotonic() + 5
        while client._process.poll() is None and time.monotonic() < deadline:
            time.sleep(0.01)
        for thread in client._threads:
            thread.join(timeout=1)
        with self.assertRaises(SessionClientError) as caught:
            client.call(self.address("writeContent"), {})
        self.assertFalse(caught.exception.may_have_effect)
        self.assertFalse(self.events.exists())
        self.assertEqual("idle_timeout", client.session_outcome["reason"])

    def test_invalid_address_and_nonfinite_params_are_rejected_before_dispatch(self):
        with self.client() as client:
            with self.assertRaises(ValueError):
                client.call({"app": "ppt", "action": "writeContent"}, {})
            with self.assertRaises(ValueError):
                client.call(self.address("writeContent"), {"value": float("nan")})
            self.assertFalse(self.events.exists())


if __name__ == "__main__":
    unittest.main()
