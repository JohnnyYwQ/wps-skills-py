import json
from pathlib import Path
import tempfile
import unittest

from wps_skills.core.trace_journal import JsonlTraceJournal


class JsonlTraceJournalTests(unittest.TestCase):
    def test_writes_separate_session_and_action_timing_journals(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            journal = JsonlTraceJournal(
                Path(temporary_directory),
                timestamp_factory=lambda: "2026-09-04T08:00:00.000Z",
                trace_id_factory=lambda: "trace-1",
            )

            session_path = journal.session_log(
                application="word",
                session_id="session-1",
            )
            journal.session_event(
                application="word",
                session_id="session-1",
                event="session.ready",
            )
            trace = journal.action_trace(
                application="word",
                session_id="session-1",
            )
            trace.event(
                "action.finished",
                elapsedMs=1250,
                address={"app": "word", "action": "inspectDocument"},
            )

            session_rows = [
                json.loads(line)
                for line in session_path.read_text(encoding="utf-8").splitlines()
            ]
            action_rows = [
                json.loads(line)
                for line in trace.trace_log.read_text(encoding="utf-8").splitlines()
            ]

        self.assertEqual("session.ready", session_rows[0]["event"])
        self.assertEqual("2026-09-04T08:00:00.000Z", session_rows[0]["ts"])
        self.assertEqual("trace-1", trace.trace_id)
        self.assertEqual("action.finished", action_rows[0]["event"])
        self.assertEqual(1250, action_rows[0]["elapsedMs"])
        self.assertNotIn("params", action_rows[0])

    def test_unavailable_journal_keeps_stable_null_trace_path(self):
        journal = JsonlTraceJournal(None, trace_id_factory=lambda: "trace-1")

        trace = journal.action_trace(
            application="word",
            session_id="session-1",
        )
        trace.event("action.finished", elapsedMs=1)

        self.assertEqual("trace-1", trace.trace_id)
        self.assertIsNone(trace.trace_log)


if __name__ == "__main__":
    unittest.main()
