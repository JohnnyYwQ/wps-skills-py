"""Real SessionHost with fake document operations and explicit channel faults."""

import json
from pathlib import Path
import sys

from wps_skills.core.action_session import (
    ActionError, ActionResponse, ActionTurn, CleanupOutcome,
    CleanupReport, DocumentResourceCleanup,
)
from wps_skills.host.session_host import SessionHost


mode = sys.argv[1]
events = Path(sys.argv[2])


class Session:
    def execute(self, request, trace):
        with events.open("a", encoding="utf-8") as stream:
            stream.write(request.address.action + "\n")
        action = request.address.action
        outcome = "unknown" if action == "terminal" else "failed" if action == "fail" else "succeeded"
        return ActionTurn(
            response=ActionResponse(
                outcome=outcome, address=request.address, session_id="fixture-session", trace=trace,
                data={"echo": dict(request.params)} if outcome == "succeeded" else None,
                error=ActionError(code="CONTENT_VERIFICATION_FAILED", message="actual font read-back error") if outcome != "succeeded" else None,
            ),
            continuation="terminate" if action == "terminal" else "continue",
        )

    def close(self):
        return CleanupOutcome(
            outcome="failed" if mode == "cleanup_failed" else "succeeded",
            cleanup=CleanupReport(document_resources=DocumentResourceCleanup(state="released")),
            error=ActionError(code="SESSION_CLEANUP_INCOMPLETE", message="fixture cleanup failure") if mode == "cleanup_failed" else None,
        )


if mode in {"eof", "mismatch", "missing_closed", "idle", "startup", "timeout"}:
    if mode == "startup":
        sys.stderr.write("fixture startup failed\n")
        raise SystemExit(4)
    print(json.dumps({"type": "session.ready", "protocolVersion": 1, "app": "word", "sessionId": "fixture-session"}), flush=True)
    if mode == "idle":
        print(json.dumps({"type": "session.closed", "session": {
            "sessionId": "fixture-session", "app": "word", "outcome": "succeeded", "reason": "idle_timeout",
        }}), flush=True)
        raise SystemExit(0)
    request = json.loads(sys.stdin.readline())
    events.write_text(request["address"]["action"] + "\n")
    if mode == "timeout":
        sys.stdin.read()
        raise SystemExit(4)
    if mode == "eof":
        raise SystemExit(4)
    if mode == "mismatch":
        print(json.dumps({"outcome": "succeeded", "address": request["address"], "sessionId": "wrong-session", "traceId": "fixture-trace", "data": {}}), flush=True)
        raise SystemExit(0)
    print(json.dumps({"type": "session.closing", "sessionId": "fixture-session", "reason": "host_failure"}), flush=True)
    print(json.dumps({"outcome": "unknown", "address": request["address"], "sessionId": "fixture-session", "traceId": "fixture-trace", "error": {"code": "CONTENT_VERIFICATION_FAILED", "message": "actual font read-back error"}}), flush=True)
    raise SystemExit(4)

if mode == "stderr":
    for _ in range(4000):
        sys.stderr.write("fixture diagnostic output\n")

host = SessionHost(session_factory=lambda **kwargs: Session(), session_id_factory=lambda: "fixture-session")
raise SystemExit(host.serve(application="word", input_stream=sys.stdin, output_stream=sys.stdout, error_stream=sys.stderr))
