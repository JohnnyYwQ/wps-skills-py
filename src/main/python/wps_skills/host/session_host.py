"""Canonical JSONL Session Host for one application-scoped Action Session."""

import json
import math
import os
import queue
import threading
import time
from typing import Any, Callable, Optional, TextIO
import uuid

from wps_skills.core.action_session import ActionAddress, ActionRequest, TraceContext


class SessionStartupError(RuntimeError):
    """The Action Session could not be transactionally constructed."""


class _DuplicateJsonKey(ValueError):
    pass


class _NonFiniteJsonNumber(ValueError):
    pass


def _reject_nonfinite_number(value):
    raise _NonFiniteJsonNumber(value)


def _parse_finite_float(value):
    parsed = float(value)
    if not math.isfinite(parsed):
        raise _NonFiniteJsonNumber(value)
    return parsed


_END_OF_INPUT = object()


class _ThreadedInputReader:
    """Read without dispatching so WAITING can enforce its idle budget."""

    def __init__(self, stream):
        self._stream = stream
        self._items = queue.Queue()
        self.ended = threading.Event()
        self._thread = threading.Thread(target=self._read, daemon=True)
        self._thread.start()

    def _read(self):
        try:
            for line in self._stream:
                self._items.put(line)
        except BaseException as exc:
            self._items.put(exc)
        finally:
            self.ended.set()
            self._items.put(_END_OF_INPUT)

    def has_buffered_line(self):
        with self._items.mutex:
            return any(
                isinstance(item, str) and bool(item.strip())
                for item in self._items.queue
            )

    def read(self, timeout_seconds):
        try:
            item = self._items.get(timeout=max(0, timeout_seconds))
        except queue.Empty as exc:
            raise TimeoutError from exc
        if item is _END_OF_INPUT:
            return None
        if isinstance(item, BaseException):
            raise item
        return item


class _SessionCloser:
    def __init__(self, session, clock):
        self._session = session
        self._clock = clock
        self._lock = threading.Lock()
        self._outcome = None
        self._elapsed_ms = 0

    def close(self):
        with self._lock:
            if self._outcome is None:
                started = self._clock()
                self._outcome = self._session.close()
                self._elapsed_ms = max(
                    0,
                    int(round((self._clock() - started) * 1000)),
                )
            return self._outcome

    @property
    def elapsed_ms(self):
        return self._elapsed_ms


class _SessionMetrics:
    def __init__(self, clock):
        self._clock = clock
        self._started_at = clock()
        self.action_count = 0
        self.action_execution_elapsed_ms = 0

    @staticmethod
    def _milliseconds(elapsed_seconds):
        return max(0, int(round(elapsed_seconds * 1000)))

    def record_action(self, started_at):
        elapsed_ms = self._milliseconds(self._clock() - started_at)
        self.action_count += 1
        self.action_execution_elapsed_ms += elapsed_ms
        return elapsed_ms

    def session_elapsed_ms(self):
        return self._milliseconds(self._clock() - self._started_at)


def _reject_duplicate_keys(pairs):
    value = {}
    for key, item in pairs:
        if key in value:
            raise _DuplicateJsonKey(key)
        value[key] = item
    return value


def _request_problem(value):
    if set(value) != {"address", "params"} or not isinstance(
        value.get("params"), dict
    ):
        return (
            "INVALID_ACTION_REQUEST",
            "Action Request must contain exactly address and object params",
        )
    address = value.get("address")
    if (
        not isinstance(address, dict)
        or set(address) != {"app", "action"}
        or address.get("app") not in {"excel", "ppt", "word"}
        or not isinstance(address.get("action"), str)
        or not address["action"]
    ):
        return (
            "INVALID_ACTION_ADDRESS",
            "address must contain canonical app and action values",
        )
    return None


class SessionHost:
    def __init__(
        self,
        *,
        session_factory: Callable[..., Any],
        session_id_factory: Callable[[], str],
        pid_factory: Callable[[], int] = os.getpid,
        trace_log_factory: Optional[Callable[..., Any]] = None,
        action_trace_factory: Optional[Callable[..., TraceContext]] = None,
        session_event_sink: Optional[Callable[..., Any]] = None,
        input_reader_factory: Optional[Callable[..., Any]] = None,
        clock: Callable[[], float] = time.monotonic,
        idle_timeout_seconds: int = 300,
        host_exit_timeout_seconds: float = 5,
    ):
        self._session_factory = session_factory
        self._session_id_factory = session_id_factory
        self._pid_factory = pid_factory
        self._trace_log_factory = trace_log_factory or (lambda **kwargs: None)
        self._action_trace_factory = action_trace_factory or (
            lambda **kwargs: TraceContext(
                trace_id=uuid.uuid4().hex,
                trace_log=None,
            )
        )
        self._session_event_sink = session_event_sink
        self._input_reader_factory = (
            input_reader_factory or _ThreadedInputReader
        )
        self._clock = clock
        self._idle_timeout_seconds = idle_timeout_seconds
        self._host_exit_timeout_seconds = host_exit_timeout_seconds

    @staticmethod
    def _write_record(stream: TextIO, record: Any) -> None:
        encoded = (
            json.dumps(
                record,
                ensure_ascii=True,
                allow_nan=False,
                separators=(",", ":"),
            )
            + "\n"
        )
        written = stream.write(encoded)
        if written is not None and written != len(encoded):
            raise OSError("partial Session Protocol write")
        stream.flush()

    def _reject(
        self,
        *,
        output_stream: TextIO,
        application: str,
        session_id: str,
        code: str,
        message: str,
    ) -> None:
        trace = self._action_trace_factory(
            application=application,
            session_id=session_id,
        )
        trace.event(
            "request.rejected",
            code=code,
        )
        self._write_record(output_stream, {
            "type": "request.rejected",
            "error": {"code": code, "message": message},
            "sessionId": session_id,
            "traceId": trace.trace_id,
            "traceLog": (
                str(trace.trace_log) if trace.trace_log is not None else None
            ),
        })

    def _safe_reject(self, *, closer, **kwargs):
        try:
            self._reject(**kwargs)
        except Exception:
            return self._cleanup_exit(
                closer.close(),
                host_failure=True,
            )
        return None

    @staticmethod
    def _cleanup_exit(cleanup, *, host_failure):
        if host_failure:
            return 4 if cleanup.outcome == "succeeded" else 6
        return 0 if cleanup.outcome == "succeeded" else 2

    def _session_event(
        self,
        *,
        application,
        session_id,
        event,
        **fields,
    ):
        if self._session_event_sink is None:
            return
        try:
            self._session_event_sink(
                application=application,
                session_id=session_id,
                event=event,
                **fields,
            )
        except Exception:
            # Session tracing is diagnostic and cannot affect protocol output.
            pass

    @staticmethod
    def _session_outcome(
        *,
        cleanup,
        session_id,
        application,
        reason,
        trace_log,
        elapsed_ms,
    ):
        outcome = {
            "outcome": cleanup.outcome,
            "sessionId": session_id,
            "app": application,
            "reason": reason,
            "cleanup": {
                "elapsedMs": elapsed_ms,
                "documentResources": {
                    "state": cleanup.cleanup.document_resources.state,
                },
                "processes": [
                    {
                        "pid": process.pid,
                        "cleanupSteps": list(process.cleanup_steps),
                        "released": process.released,
                    }
                    for process in cleanup.cleanup.processes
                ],
            },
            "traceLog": trace_log,
        }
        if cleanup.error is not None:
            outcome["error"] = {
                "code": cleanup.error.code,
                "message": cleanup.error.message,
            }
        return outcome

    def _finish(
        self,
        *,
        closer,
        output_stream,
        session_id,
        application,
        reason,
        trace_log,
        metrics,
    ):
        cleanup = closer.close()
        session_elapsed_ms = metrics.session_elapsed_ms()
        self._session_event(
            application=application,
            session_id=session_id,
            event="session.finished",
            reason=reason,
            outcome=cleanup.outcome,
            sessionElapsedMs=session_elapsed_ms,
            actionCount=metrics.action_count,
            actionExecutionElapsedMs=(
                metrics.action_execution_elapsed_ms
            ),
            cleanupElapsedMs=closer.elapsed_ms,
        )
        record = {
            "type": "session.closed",
            "session": self._session_outcome(
                cleanup=cleanup,
                session_id=session_id,
                application=application,
                reason=reason,
                trace_log=trace_log,
                elapsed_ms=closer.elapsed_ms,
            ),
        }
        try:
            self._write_record(output_stream, record)
        except Exception:
            return self._cleanup_exit(cleanup, host_failure=True)
        return self._cleanup_exit(
            cleanup,
            host_failure=reason == "host_failure",
        )

    @staticmethod
    def _disconnect_monitor(reader, closer):
        if not hasattr(reader, "ended"):
            return None
        stop = threading.Event()
        lost = threading.Event()

        def monitor():
            while not stop.is_set():
                if reader.ended.wait(timeout=0.01):
                    if (
                        not stop.is_set()
                        and not reader.has_buffered_line()
                    ):
                        lost.set()
                        closer.close()
                    return

        worker = threading.Thread(target=monitor, daemon=True)
        worker.start()
        return stop, lost

    def _execute_action(
        self,
        *,
        session,
        request,
        trace,
        monitor,
    ):
        completed = threading.Event()
        result = {}

        def execute():
            try:
                result["turn"] = session.execute(request, trace)
            except BaseException as exc:
                result["error"] = exc
            finally:
                completed.set()

        worker = threading.Thread(target=execute, daemon=True)
        worker.start()
        while not completed.wait(timeout=0.01):
            if monitor is not None and monitor[1].is_set():
                if not completed.wait(
                    timeout=self._host_exit_timeout_seconds
                ):
                    return "channel_lost", None
                break
        if monitor is not None:
            monitor[0].set()
            if monitor[1].is_set():
                return "channel_lost", None
        if "error" in result:
            return "error", result["error"]
        return "turn", result["turn"]

    def serve(
        self,
        *,
        application: str,
        input_stream: TextIO,
        output_stream: TextIO,
        error_stream: TextIO,
    ) -> int:
        session_id = self._session_id_factory()
        try:
            trace_log = self._trace_log_factory(
                application=application,
                session_id=session_id,
            )
            trace_log = str(trace_log) if trace_log is not None else None
            session = self._session_factory(
                application=application,
                session_id=session_id,
            )
        except Exception as exc:
            error_stream.write(
                f"WPS_SESSION_STARTUP_FAILED app={application}: {exc}\n"
            )
            error_stream.flush()
            return 4
        closer = _SessionCloser(session, self._clock)
        metrics = _SessionMetrics(self._clock)
        pid = self._pid_factory()
        try:
            self._write_record(output_stream, {
                "type": "session.ready",
                "protocolVersion": 1,
                "sessionId": session_id,
                "app": application,
                "pid": pid,
                "idleTimeoutSeconds": self._idle_timeout_seconds,
                "traceLog": trace_log,
            })
        except Exception:
            cleanup = closer.close()
            return self._cleanup_exit(cleanup, host_failure=True)
        self._session_event(
            application=application,
            session_id=session_id,
            event="session.ready",
            protocolVersion=1,
            pid=pid,
        )
        reader = self._input_reader_factory(input_stream)
        waiting_since = self._clock()
        while True:
            remaining = self._idle_timeout_seconds - (
                self._clock() - waiting_since
            )
            try:
                line = reader.read(remaining)
            except TimeoutError:
                return self._finish(
                    closer=closer,
                    output_stream=output_stream,
                    session_id=session_id,
                    application=application,
                    reason="idle_timeout",
                    trace_log=trace_log,
                    metrics=metrics,
                )
            except Exception:
                cleanup = closer.close()
                return self._cleanup_exit(cleanup, host_failure=True)
            if line is None:
                cleanup = closer.close()
                return self._cleanup_exit(cleanup, host_failure=True)
            if not line.strip():
                continue
            try:
                value = json.loads(
                    line,
                    object_pairs_hook=_reject_duplicate_keys,
                    parse_constant=_reject_nonfinite_number,
                    parse_float=_parse_finite_float,
                )
            except (
                json.JSONDecodeError,
                _DuplicateJsonKey,
                _NonFiniteJsonNumber,
            ):
                failed_exit = self._safe_reject(
                    closer=closer,
                    output_stream=output_stream,
                    application=application,
                    session_id=session_id,
                    code="INVALID_ACTION_REQUEST",
                    message="Input must be one valid JSON object",
                )
                if failed_exit is not None:
                    return failed_exit
                waiting_since = self._clock()
                continue
            if value == {"control": "close"}:
                return self._finish(
                    closer=closer,
                    output_stream=output_stream,
                    session_id=session_id,
                    application=application,
                    reason="client_close",
                    trace_log=trace_log,
                    metrics=metrics,
                )
            if not isinstance(value, dict):
                failed_exit = self._safe_reject(
                    closer=closer,
                    output_stream=output_stream,
                    application=application,
                    session_id=session_id,
                    code="INVALID_ACTION_REQUEST",
                    message="Input must be one valid JSON object",
                )
                if failed_exit is not None:
                    return failed_exit
                waiting_since = self._clock()
                continue
            problem = _request_problem(value)
            if problem is not None:
                failed_exit = self._safe_reject(
                    closer=closer,
                    output_stream=output_stream,
                    application=application,
                    session_id=session_id,
                    code=problem[0],
                    message=problem[1],
                )
                if failed_exit is not None:
                    return failed_exit
                waiting_since = self._clock()
                continue
            trace = self._action_trace_factory(
                application=application,
                session_id=session_id,
            )
            request = ActionRequest(
                address=ActionAddress(
                    app=value["address"]["app"],
                    action=value["address"]["action"],
                ),
                params=value["params"],
            )
            monitor = self._disconnect_monitor(reader, closer)
            action_started_at = self._clock()
            address = {
                "app": request.address.app,
                "action": request.address.action,
            }
            trace.event("action.started", address=address)
            execution, value = self._execute_action(
                session=session,
                request=request,
                trace=trace,
                monitor=monitor,
            )
            action_elapsed_ms = metrics.record_action(action_started_at)
            action_outcome = (
                value.response.outcome
                if execution == "turn" and value is not None
                else None
            )
            trace.event(
                "action.finished",
                address=address,
                execution=execution,
                outcome=action_outcome,
                elapsedMs=action_elapsed_ms,
            )
            self._session_event(
                application=application,
                session_id=session_id,
                event="action.finished",
                traceId=trace.trace_id,
                address=address,
                execution=execution,
                outcome=action_outcome,
                elapsedMs=action_elapsed_ms,
            )
            if execution == "channel_lost":
                cleanup = closer.close()
                return self._cleanup_exit(
                    cleanup,
                    host_failure=True,
                )
            if execution == "error":
                try:
                    self._write_record(output_stream, {
                        "type": "session.closing",
                        "sessionId": session_id,
                        "reason": "host_failure",
                    })
                except Exception:
                    cleanup = closer.close()
                    return self._cleanup_exit(
                        cleanup,
                        host_failure=True,
                    )
                return self._finish(
                    closer=closer,
                    output_stream=output_stream,
                    session_id=session_id,
                    application=application,
                    reason="host_failure",
                    trace_log=trace_log,
                    metrics=metrics,
                )
            turn = value
            if turn.continuation == "terminate":
                try:
                    self._write_record(output_stream, {
                        "type": "session.closing",
                        "sessionId": session_id,
                        "reason": "host_failure",
                    })
                    self._write_record(
                        output_stream,
                        turn.response.to_wire(),
                    )
                except Exception:
                    cleanup = closer.close()
                    return self._cleanup_exit(
                        cleanup,
                        host_failure=True,
                    )
                return self._finish(
                    closer=closer,
                    output_stream=output_stream,
                    session_id=session_id,
                    application=application,
                    reason="host_failure",
                    trace_log=trace_log,
                    metrics=metrics,
                )
            try:
                self._write_record(output_stream, turn.response.to_wire())
            except Exception:
                cleanup = closer.close()
                return self._cleanup_exit(cleanup, host_failure=True)
            waiting_since = self._clock()
