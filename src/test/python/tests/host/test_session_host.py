import io
import json
from pathlib import Path
import threading
import unittest

from wps_skills.core.action_session import (
    ActionAddress,
    ActionError,
    ActionRequest,
    ActionResponse,
    ActionTurn,
    CleanupOutcome,
    CleanupReport,
    DocumentResourceCleanup,
    ProcessCleanup,
    TraceContext,
)
from wps_skills.host.session_host import SessionHost, SessionStartupError


class FakeActionSession:
    def __init__(
        self,
        *,
        cleanup_state="not_acquired",
        cleanup_outcome="succeeded",
        continuation="continue",
        action_error_code="SESSION_DOCUMENT_NOT_BOUND",
        execute_error=None,
    ):
        self.cleanup_state = cleanup_state
        self.cleanup_outcome = cleanup_outcome
        self.continuation = continuation
        self.action_error_code = action_error_code
        self.execute_error = execute_error
        self.close_count = 0
        self.execute_calls = []

    def execute(self, request, trace):
        self.execute_calls.append((request, trace))
        if self.execute_error is not None:
            raise self.execute_error
        return ActionTurn(
            response=ActionResponse(
                outcome="failed",
                address=request.address,
                session_id="session-1",
                trace=trace,
                error=ActionError(
                    code=self.action_error_code,
                    message="The Action Session cannot continue",
                ),
            ),
            continuation=self.continuation,
        )

    def close(self):
        self.close_count += 1
        return CleanupOutcome(
            outcome=self.cleanup_outcome,
            cleanup=CleanupReport(
                document_resources=DocumentResourceCleanup(
                    state=self.cleanup_state,
                ),
            ),
            error=(
                ActionError(
                    code="SESSION_CLEANUP_INCOMPLETE",
                    message="Cleanup could not be confirmed",
                )
                if self.cleanup_outcome == "failed"
                else None
            ),
        )


class FailingOutput(io.StringIO):
    def write(self, value):
        raise OSError("output channel closed")


class FailingNthWrite(io.StringIO):
    def __init__(self, fail_at):
        super().__init__()
        self.fail_at = fail_at
        self.write_count = 0

    def write(self, value):
        self.write_count += 1
        if self.write_count == self.fail_at:
            raise OSError("output channel closed")
        return super().write(value)


class TimingOutReader:
    def read(self, timeout_seconds):
        raise TimeoutError


class ControlledInput:
    def __init__(self, first_line):
        self.first_line = first_line
        self.allow_eof = threading.Event()

    def __iter__(self):
        yield self.first_line
        self.allow_eof.wait(timeout=2)


class ManualClock:
    def __init__(self):
        self.now = 0.0

    def __call__(self):
        return self.now

    def advance(self, seconds):
        self.now += seconds


class TimedActionSession(FakeActionSession):
    def __init__(self, clock):
        super().__init__()
        self.clock = clock

    def execute(self, request, trace):
        self.clock.advance(1.25)
        return super().execute(request, trace)

    def close(self):
        self.clock.advance(0.2)
        return super().close()


class SessionHostTests(unittest.TestCase):
    def test_protocol_output_is_ascii_safe_json(self):
        output = io.StringIO()

        SessionHost._write_record(output, {"value": "中文"})

        self.assertEqual({"value": "中文"}, json.loads(output.getvalue()))
        self.assertIn(r"\u4e2d\u6587", output.getvalue())
        self.assertNotIn("中文", output.getvalue())

    def test_session_outcome_serializes_owned_process_cleanup(self):
        cleanup = CleanupOutcome(
            outcome="succeeded",
            cleanup=CleanupReport(
                document_resources=DocumentResourceCleanup(
                    state="released",
                ),
                processes=(ProcessCleanup(
                    pid=42,
                    cleanup_steps=("graceful",),
                    released=True,
                ),),
            ),
        )

        outcome = SessionHost._session_outcome(
            cleanup=cleanup,
            session_id="session-1",
            application="word",
            reason="client_close",
            trace_log=None,
            elapsed_ms=7,
        )

        self.assertEqual(7, outcome["cleanup"]["elapsedMs"])
        self.assertEqual(
            [{
                "pid": 42,
                "cleanupSteps": ["graceful"],
                "released": True,
            }],
            outcome["cleanup"]["processes"],
        )

    def test_missing_production_application_slice_fails_before_ready(self):
        calls = []

        def unavailable_factory(*, application, session_id):
            calls.append((application, session_id))
            raise SessionStartupError(
                "no production Application Contract Set and Adapter are installed"
            )

        output = io.StringIO()
        errors = io.StringIO()
        host = SessionHost(
            session_factory=unavailable_factory,
            session_id_factory=lambda: "session-1",
        )

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO(""),
            output_stream=output,
            error_stream=errors,
        )

        self.assertEqual(4, exit_code)
        self.assertEqual("", output.getvalue())
        self.assertIn("WPS_SESSION_STARTUP_FAILED", errors.getvalue())
        self.assertIn("app=word", errors.getvalue())
        self.assertIn("no production Application Contract Set", errors.getvalue())
        self.assertEqual([("word", "session-1")], calls)

    def test_ready_write_failure_closes_constructed_session_and_returns_host_failure(self):
        session = FakeActionSession()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
        )

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO(""),
            output_stream=FailingOutput(),
            error_stream=io.StringIO(),
        )

        self.assertEqual(4, exit_code)
        self.assertEqual(1, session.close_count)

    def test_action_response_write_failure_closes_session_as_channel_loss(self):
        session = FakeActionSession()
        output = FailingNthWrite(fail_at=2)
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
            action_trace_factory=lambda **kwargs: TraceContext(
                trace_id="trace-1",
                trace_log=None,
            ),
        )
        request = json.dumps({
            "address": {"app": "word", "action": "inspectDocument"},
            "params": {},
        }) + "\n"
        inputs = ControlledInput(request)

        exit_code = host.serve(
            application="word",
            input_stream=inputs,
            output_stream=output,
            error_stream=io.StringIO(),
        )

        self.assertEqual(4, exit_code)
        self.assertEqual(1, session.close_count)
        self.assertEqual(1, len(session.execute_calls))

    def test_rejection_write_failure_closes_session_as_channel_loss(self):
        session = FakeActionSession()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
            action_trace_factory=lambda **kwargs: TraceContext(
                trace_id="trace-1",
                trace_log=None,
            ),
        )

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO("{\n"),
            output_stream=FailingNthWrite(fail_at=2),
            error_stream=io.StringIO(),
        )

        self.assertEqual(4, exit_code)
        self.assertEqual(1, session.close_count)
        self.assertEqual([], session.execute_calls)

    def test_eof_cleanup_failure_sets_both_host_and_cleanup_exit_bits(self):
        session = FakeActionSession(
            cleanup_state="release_unconfirmed",
            cleanup_outcome="failed",
        )
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
        )

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO(""),
            output_stream=output,
            error_stream=io.StringIO(),
        )

        self.assertEqual(6, exit_code)
        self.assertEqual(1, session.close_count)
        self.assertEqual(1, len(output.getvalue().splitlines()))

    def test_waiting_idle_timeout_closes_normally_without_closing_signal(self):
        session = FakeActionSession()
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
            input_reader_factory=lambda stream: TimingOutReader(),
        )

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO(""),
            output_stream=output,
            error_stream=io.StringIO(),
        )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertEqual(0, exit_code)
        self.assertEqual(
            ["session.ready", "session.closed"],
            [record["type"] for record in records],
        )
        self.assertEqual("idle_timeout", records[1]["session"]["reason"])
        self.assertEqual(1, session.close_count)

    def test_closed_record_write_failure_sets_transport_exit_bit_without_reclosing(self):
        session = FakeActionSession()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
        )

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO('{"control":"close"}\n'),
            output_stream=FailingNthWrite(fail_at=2),
            error_stream=io.StringIO(),
        )

        self.assertEqual(4, exit_code)
        self.assertEqual(1, session.close_count)

    def test_ready_then_structured_close_emits_successful_session_outcome(self):
        session = FakeActionSession()
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
        )

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO('{"control":"close"}\n'),
            output_stream=output,
            error_stream=io.StringIO(),
        )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertEqual(0, exit_code)
        self.assertEqual(
            {
                "type": "session.ready",
                "protocolVersion": 1,
                "sessionId": "session-1",
                "app": "word",
                "pid": 123,
                "idleTimeoutSeconds": 300,
                "traceLog": None,
            },
            records[0],
        )
        self.assertEqual("session.closed", records[1]["type"])
        self.assertEqual(
            {
                "outcome": "succeeded",
                "sessionId": "session-1",
                "app": "word",
                "reason": "client_close",
                "cleanup": {
                    "elapsedMs": 0,
                    "documentResources": {"state": "not_acquired"},
                    "processes": [],
                },
                "traceLog": None,
            },
            records[1]["session"],
        )
        self.assertEqual(1, session.close_count)

    def test_trace_log_path_is_serialized_as_a_wire_string(self):
        session = FakeActionSession()
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: Path("logs/session-1.jsonl"),
        )

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO('{"control":"close"}\n'),
            output_stream=output,
            error_stream=io.StringIO(),
        )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertEqual(0, exit_code)
        self.assertEqual("logs/session-1.jsonl", records[0]["traceLog"])
        self.assertEqual(
            "logs/session-1.jsonl",
            records[1]["session"]["traceLog"],
        )

    def test_action_and_session_timing_stays_in_diagnostic_traces(self):
        clock = ManualClock()
        session = TimedActionSession(clock)
        action_events = []
        session_events = []
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: Path("logs/session-1.jsonl"),
            action_trace_factory=lambda **kwargs: TraceContext(
                trace_id="trace-1",
                trace_log=Path("logs/trace-1.jsonl"),
                event_sink=lambda event, **fields: action_events.append(
                    (event, fields)
                ),
            ),
            session_event_sink=lambda **record: session_events.append(record),
            clock=clock,
        )
        inputs = "\n".join([
            json.dumps({
                "address": {"app": "word", "action": "inspectDocument"},
                "params": {},
            }),
            json.dumps({"control": "close"}),
            "",
        ])

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO(inputs),
            output_stream=output,
            error_stream=io.StringIO(),
        )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertEqual(0, exit_code)
        self.assertNotIn("elapsedMs", records[1])
        self.assertEqual(
            ["action.started", "action.finished"],
            [event for event, fields in action_events],
        )
        self.assertEqual(1250, action_events[-1][1]["elapsedMs"])
        self.assertEqual(
            ["session.ready", "action.finished", "session.finished"],
            [record["event"] for record in session_events],
        )
        summary = session_events[-1]
        self.assertEqual(1, summary["actionCount"])
        self.assertEqual(1250, summary["actionExecutionElapsedMs"])
        self.assertEqual(200, summary["cleanupElapsedMs"])
        self.assertEqual(1450, summary["sessionElapsedMs"])

    def test_canonical_action_request_receives_exact_action_response(self):
        session = FakeActionSession()
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
            action_trace_factory=lambda **kwargs: TraceContext(
                trace_id="trace-1",
                trace_log=None,
            ),
        )
        inputs = "\n".join([
            json.dumps({
                "address": {"app": "word", "action": "inspectDocument"},
                "params": {},
            }),
            json.dumps({"control": "close"}),
            "",
        ])

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO(inputs),
            output_stream=output,
            error_stream=io.StringIO(),
        )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertEqual(0, exit_code)
        self.assertEqual(
            {
                "outcome": "failed",
                "address": {"app": "word", "action": "inspectDocument"},
                "error": {
                    "code": "SESSION_DOCUMENT_NOT_BOUND",
                    "message": "The Action Session cannot continue",
                },
                "sessionId": "session-1",
                "traceId": "trace-1",
                "traceLog": None,
            },
            records[1],
        )
        self.assertNotIn("success", records[1])
        self.assertEqual(
            [(
                ActionRequest(
                    address=ActionAddress(
                        app="word",
                        action="inspectDocument",
                    ),
                    params={},
                ),
                TraceContext(trace_id="trace-1", trace_log=None),
            )],
            session.execute_calls,
        )

    def test_invalid_json_duplicate_keys_and_bare_exit_are_traced_rejections(self):
        session = FakeActionSession()
        trace_ids = iter(["trace-1", "trace-2", "trace-3", "trace-4"])
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
            action_trace_factory=lambda **kwargs: TraceContext(
                trace_id=next(trace_ids),
                trace_log=None,
            ),
        )
        inputs = "\n".join([
            "{",
            '{"address":{"app":"word","action":"inspectDocument"},'
            '"params":{},"params":{}}',
            '"EXIT"',
            '{"address":{"app":"word","action":"inspectDocument"},'
            '"params":{"value":NaN}}',
            '{"control":"close"}',
            "",
        ])

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO(inputs),
            output_stream=output,
            error_stream=io.StringIO(),
        )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        rejections = records[1:5]
        self.assertEqual(0, exit_code)
        self.assertEqual(
            ["request.rejected"] * 4,
            [record["type"] for record in rejections],
        )
        self.assertEqual(
            ["INVALID_ACTION_REQUEST"] * 4,
            [record["error"]["code"] for record in rejections],
        )
        self.assertEqual(
            ["trace-1", "trace-2", "trace-3", "trace-4"],
            [record["traceId"] for record in rejections],
        )
        self.assertEqual([], session.execute_calls)

    def test_terminal_turn_closes_before_response_and_ignores_buffered_input(self):
        session = FakeActionSession(
            continuation="terminate",
            action_error_code="DOCUMENT_CLOSED",
        )
        trace_ids = iter(["trace-1", "trace-2"])
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
            action_trace_factory=lambda **kwargs: TraceContext(
                trace_id=next(trace_ids),
                trace_log=None,
            ),
        )
        request = json.dumps({
            "address": {"app": "word", "action": "inspectDocument"},
            "params": {},
        })

        exit_code = host.serve(
            application="word",
            input_stream=io.StringIO(f"{request}\n{request}\n"),
            output_stream=output,
            error_stream=io.StringIO(),
        )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertEqual(4, exit_code)
        self.assertEqual("session.ready", records[0]["type"])
        self.assertEqual("session.closing", records[1]["type"])
        self.assertEqual("host_failure", records[1]["reason"])
        self.assertEqual("DOCUMENT_CLOSED", records[2]["error"]["code"])
        self.assertEqual("session.closed", records[3]["type"])
        self.assertEqual("host_failure", records[3]["session"]["reason"])
        self.assertEqual(1, len(session.execute_calls))
        self.assertEqual(1, session.close_count)

    def test_unexpected_session_exception_emits_host_failure_lifecycle(self):
        session = FakeActionSession(
            execute_error=RuntimeError("unexpected core failure"),
        )
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
            action_trace_factory=lambda **kwargs: TraceContext(
                trace_id="trace-1",
                trace_log=None,
            ),
        )
        request = json.dumps({
            "address": {"app": "word", "action": "inspectDocument"},
            "params": {},
        }) + "\n"
        inputs = ControlledInput(request)

        exit_code = host.serve(
            application="word",
            input_stream=inputs,
            output_stream=output,
            error_stream=io.StringIO(),
        )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertEqual(4, exit_code)
        self.assertEqual(
            ["session.ready", "session.closing", "session.closed"],
            [record["type"] for record in records],
        )
        self.assertEqual("host_failure", records[1]["reason"])
        self.assertEqual("host_failure", records[2]["session"]["reason"])
        self.assertEqual(1, session.close_count)

    def test_eof_during_blocked_action_triggers_independent_cleanup(self):
        entered = threading.Event()
        resume = threading.Event()
        closed = threading.Event()

        class BlockingSession(FakeActionSession):
            def execute(self, request, trace):
                self.execute_calls.append((request, trace))
                entered.set()
                resume.wait(timeout=2)
                return super().execute(request, trace)

            def close(self):
                outcome = super().close()
                closed.set()
                return outcome

        session = BlockingSession()
        request = json.dumps({
            "address": {"app": "word", "action": "inspectDocument"},
            "params": {},
        }) + "\n"
        inputs = ControlledInput(request)
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
            action_trace_factory=lambda **kwargs: TraceContext(
                trace_id="trace-1",
                trace_log=None,
            ),
            host_exit_timeout_seconds=0.02,
        )
        exit_codes = []
        worker = threading.Thread(target=lambda: exit_codes.append(host.serve(
            application="word",
            input_stream=inputs,
            output_stream=output,
            error_stream=io.StringIO(),
        )))
        worker.start()
        self.assertTrue(entered.wait(timeout=1))

        inputs.allow_eof.set()
        self.assertTrue(closed.wait(timeout=1))
        worker.join(timeout=1)

        self.assertFalse(worker.is_alive())
        self.assertEqual([4], exit_codes)
        self.assertEqual(1, session.close_count)
        self.assertEqual(1, len(output.getvalue().splitlines()))
        resume.set()
        self.assertEqual(1, session.close_count)

    def test_request_and_address_shapes_are_rejected_before_session_execution(self):
        session = FakeActionSession()
        trace_ids = iter(["trace-1", "trace-2"])
        output = io.StringIO()
        host = SessionHost(
            session_factory=lambda **kwargs: session,
            session_id_factory=lambda: "session-1",
            pid_factory=lambda: 123,
            trace_log_factory=lambda **kwargs: None,
            action_trace_factory=lambda **kwargs: TraceContext(
                trace_id=next(trace_ids),
                trace_log=None,
            ),
        )
        inputs = "\n".join([
            json.dumps({
                "address": {"app": "word", "action": "inspectDocument"},
                "params": {},
                "extra": True,
            }),
            json.dumps({
                "address": {"app": "pages", "action": "inspectDocument"},
                "params": {},
            }),
            json.dumps({"control": "close"}),
            "",
        ])

        host.serve(
            application="word",
            input_stream=io.StringIO(inputs),
            output_stream=output,
            error_stream=io.StringIO(),
        )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertEqual(
            ["INVALID_ACTION_REQUEST", "INVALID_ACTION_ADDRESS"],
            [records[1]["error"]["code"], records[2]["error"]["code"]],
        )
        self.assertEqual([], session.execute_calls)


if __name__ == "__main__":
    unittest.main()
