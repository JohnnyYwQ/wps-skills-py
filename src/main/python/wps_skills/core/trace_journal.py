"""Best-effort JSONL timing traces for Action Sessions and Actions."""

from datetime import datetime, timezone
import json
import os
from pathlib import Path
import re
import tempfile
import threading
from typing import Callable, Optional
import uuid

from wps_skills.core.action_session import TraceContext


_SAFE_FILE_NAME = re.compile(r"[^A-Za-z0-9._-]+")


def _safe_file_name(value: str) -> str:
    safe = _SAFE_FILE_NAME.sub("-", value).strip("-.")
    return safe[:160] or uuid.uuid4().hex


def _utc_timestamp() -> str:
    return (
        datetime.now(timezone.utc)
        .isoformat(timespec="milliseconds")
        .replace("+00:00", "Z")
    )


def _default_roots():
    override = os.environ.get("WPS_TRACE_DIR")
    if override:
        yield Path(override).expanduser()
    if os.name == "nt":
        local_app_data = os.environ.get("LOCALAPPDATA")
        if local_app_data:
            yield Path(local_app_data) / "wps-skills" / "logs"
    else:
        state_home = os.environ.get("XDG_STATE_HOME")
        if state_home:
            yield Path(state_home) / "wps-skills"
        try:
            yield Path.home() / ".local" / "state" / "wps-skills"
        except Exception:
            pass
    yield Path(tempfile.gettempdir()) / "wps-skills-logs"


class JsonlTraceJournal:
    """Hide trace paths, timestamping, serialization, and write degradation."""

    def __init__(
        self,
        root: Optional[Path],
        *,
        timestamp_factory: Callable[[], str] = _utc_timestamp,
        trace_id_factory: Callable[[], str] = (
            lambda: f"trace-{uuid.uuid4().hex}"
        ),
    ):
        self._root = None if root is None else Path(root)
        self._timestamp_factory = timestamp_factory
        self._trace_id_factory = trace_id_factory
        self._write_lock = threading.Lock()
        if self._root is not None:
            try:
                (self._root / "sessions").mkdir(parents=True, exist_ok=True)
                (self._root / "actions").mkdir(parents=True, exist_ok=True)
            except OSError:
                self._root = None

    @classmethod
    def default(cls):
        for root in _default_roots():
            journal = cls(root)
            if journal.available:
                return journal
        return cls(None)

    @property
    def available(self) -> bool:
        return self._root is not None

    def session_log(self, *, application: str, session_id: str):
        del application
        if self._root is None:
            return None
        return self._root / "sessions" / (
            f"{_safe_file_name(session_id)}.jsonl"
        )

    def _write(self, path: Optional[Path], row) -> bool:
        if path is None:
            return False
        record = {
            "ts": self._timestamp_factory(),
            **row,
            "pid": os.getpid(),
            "thread": threading.current_thread().name,
        }
        try:
            encoded = (
                json.dumps(
                    record,
                    ensure_ascii=False,
                    allow_nan=False,
                    separators=(",", ":"),
                )
                + "\n"
            )
            with self._write_lock:
                with path.open("a", encoding="utf-8", newline="") as stream:
                    stream.write(encoded)
                    stream.flush()
            return True
        except (OSError, TypeError, ValueError):
            return False

    def session_event(
        self,
        *,
        application: str,
        session_id: str,
        event: str,
        **fields,
    ) -> bool:
        return self._write(
            self.session_log(
                application=application,
                session_id=session_id,
            ),
            {
                "event": event,
                "component": "session_host",
                "sessionId": session_id,
                "app": application,
                **fields,
            },
        )

    def action_trace(self, *, application: str, session_id: str):
        trace_id = self._trace_id_factory()
        path = (
            None
            if self._root is None
            else self._root / "actions" / (
                f"{_safe_file_name(trace_id)}.jsonl"
            )
        )

        def emit(event: str, **fields) -> None:
            self._write(path, {
                "event": event,
                "component": "action",
                "traceId": trace_id,
                "sessionId": session_id,
                "app": application,
                **fields,
            })

        return TraceContext(
            trace_id=trace_id,
            trace_log=path,
            event_sink=emit,
        )
