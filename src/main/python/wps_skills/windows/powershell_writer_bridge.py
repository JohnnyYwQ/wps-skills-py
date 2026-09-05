"""Strict Python client for the local Windows Writer PowerShell bridge."""

import json
import math
import queue
import threading
import time
import uuid
from typing import Any, Mapping, Optional, Protocol, runtime_checkable

from wps_skills.core.action_session import ControllerContext
from wps_skills.word.adapter import WriterBackendActionFailure


@runtime_checkable
class WriterBridgeTransport(Protocol):
    """One-request-at-a-time record exchange with the owned bridge process."""

    def exchange(
        self,
        request: Mapping[str, Any],
        deadline_at: Optional[float],
    ) -> Mapping[str, Any]:
        ...

    def close(self) -> bool:
        ...


class WriterBridgeTransportError(RuntimeError):
    """The owned Writer bridge channel cannot provide a reliable response."""


class _DuplicateJsonKey(ValueError):
    pass


class _NonFiniteJsonNumber(ValueError):
    pass


def _reject_duplicate_keys(pairs):
    value = {}
    for key, item in pairs:
        if key in value:
            raise _DuplicateJsonKey(key)
        value[key] = item
    return value


def _reject_nonfinite(value):
    raise _NonFiniteJsonNumber(value)


_END_OF_OUTPUT = object()


class _BridgeOutputReader:
    def __init__(self, stream):
        self._stream = stream
        self._items = queue.Queue()
        self._thread = threading.Thread(target=self._read, daemon=True)
        self._thread.start()

    def _read(self):
        try:
            while True:
                line = self._stream.readline()
                if line == "":
                    break
                self._items.put(line)
        except BaseException as exc:
            self._items.put(exc)
        finally:
            self._items.put(_END_OF_OUTPUT)

    def read(self, timeout_seconds):
        try:
            item = self._items.get(timeout=max(0, timeout_seconds))
        except queue.Empty as exc:
            raise TimeoutError from exc
        if item is _END_OF_OUTPUT:
            return None
        if isinstance(item, BaseException):
            raise item
        return item


class JsonLineWriterBridgeTransport:
    """Serialize one strict JSON record at a time over an owned process."""

    def __init__(
        self,
        *,
        process,
        clock=time.monotonic,
        liveness_timeout_seconds=5,
        close_timeout_seconds=2,
        reader_factory=_BridgeOutputReader,
    ):
        if (
            not hasattr(process, "stdin")
            or not hasattr(process, "stdout")
            or not callable(getattr(process, "poll", None))
            or not callable(getattr(process, "wait", None))
        ):
            raise ValueError("process must expose bridge pipes and lifecycle")
        if liveness_timeout_seconds <= 0 or close_timeout_seconds <= 0:
            raise ValueError("bridge transport timeouts must be positive")
        self._process = process
        self._clock = clock
        self._liveness_timeout_seconds = liveness_timeout_seconds
        self._close_timeout_seconds = close_timeout_seconds
        self._reader = reader_factory(process.stdout)
        self._exchange_lock = threading.Lock()
        self._close_lock = threading.Lock()
        self._closed = False

    def _timeout(self, deadline_at):
        if deadline_at is None:
            return self._liveness_timeout_seconds
        if (
            not isinstance(deadline_at, (int, float))
            or isinstance(deadline_at, bool)
            or not math.isfinite(deadline_at)
        ):
            raise ValueError("bridge deadline must be a finite number")
        return max(0, deadline_at - self._clock())

    @staticmethod
    def _encode(request):
        try:
            return json.dumps(
                request,
                ensure_ascii=True,
                allow_nan=False,
                separators=(",", ":"),
            ) + "\n"
        except (TypeError, ValueError) as exc:
            raise ValueError("bridge request is not finite JSON") from exc

    @staticmethod
    def _decode(line):
        try:
            value = json.loads(
                line,
                object_pairs_hook=_reject_duplicate_keys,
                parse_constant=_reject_nonfinite,
            )
        except (json.JSONDecodeError, _DuplicateJsonKey, _NonFiniteJsonNumber) as exc:
            raise WriterBridgeTransportError(
                "Writer bridge returned invalid JSON"
            ) from exc
        if not isinstance(value, dict):
            raise WriterBridgeTransportError(
                "Writer bridge response must be a JSON object"
            )
        return value

    def exchange(self, request, deadline_at):
        encoded = self._encode(request)
        with self._exchange_lock:
            if self._closed:
                raise WriterBridgeTransportError(
                    "Writer bridge transport is closed"
                )
            if self._timeout(deadline_at) <= 0:
                raise WriterBridgeTransportError(
                    "Writer bridge deadline expired before dispatch"
                )
            if self._process.poll() is not None:
                raise WriterBridgeTransportError(
                    "Writer bridge process exited before dispatch"
                )
            try:
                written = self._process.stdin.write(encoded)
                if written is not None and written != len(encoded):
                    raise OSError("partial bridge request write")
                self._process.stdin.flush()
            except Exception as exc:
                raise WriterBridgeTransportError(
                    "Writer bridge request could not be dispatched"
                ) from exc
            try:
                line = self._reader.read(self._timeout(deadline_at))
            except TimeoutError as exc:
                raise WriterBridgeTransportError(
                    "Writer bridge response deadline expired"
                ) from exc
            except Exception as exc:
                raise WriterBridgeTransportError(
                    "Writer bridge response could not be read"
                ) from exc
            if line is None:
                raise WriterBridgeTransportError(
                    "Writer bridge exited without a response"
                )
            return self._decode(line)

    def close(self):
        with self._close_lock:
            if self._closed:
                return self._process.poll() is not None
            self._closed = True
            try:
                self._process.stdin.close()
            except Exception:
                return False
            try:
                self._process.wait(timeout=self._close_timeout_seconds)
            except Exception:
                return False
            return self._process.poll() is not None


def _plain_json(value):
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("bridge arguments require finite numbers")
        return value
    if isinstance(value, Mapping):
        plain = {}
        for key, item in value.items():
            if not isinstance(key, str):
                raise ValueError("bridge argument object keys must be strings")
            plain[key] = _plain_json(item)
        return plain
    if isinstance(value, (list, tuple)):
        return [_plain_json(item) for item in value]
    raise ValueError("bridge arguments must contain only JSON values")


def _require_mapping(value, description):
    if not isinstance(value, Mapping):
        raise TypeError(f"{description} must be an object")
    return value


def _require_exact_fields(value, fields, description):
    value = _require_mapping(value, description)
    if set(value) != set(fields):
        raise TypeError(f"{description} has invalid fields")
    return value


def _require_nonempty_string(value, description):
    if not isinstance(value, str) or not value:
        raise TypeError(f"{description} must be a non-empty string")
    return value


class PowerShellWriterBridge:
    """Translate strict bridge records into the Windows Writer bridge seam."""

    def __init__(self, *, transport: WriterBridgeTransport):
        if not isinstance(transport, WriterBridgeTransport):
            raise ValueError("transport must satisfy the Writer bridge seam")
        self._transport = transport

    def execute(self, operation, arguments, context):
        operation = _require_nonempty_string(operation, "bridge operation")
        arguments = _plain_json(
            _require_mapping(arguments, "bridge arguments")
        )
        if context is None:
            suffix = uuid.uuid4().hex
            request_id = f"liveness-{suffix}"
            trace_id = f"liveness-{suffix}"
            deadline_at = None
        elif isinstance(context, ControllerContext):
            request_id = _require_nonempty_string(
                context.request_id,
                "Controller request id",
            )
            trace_id = _require_nonempty_string(
                context.trace_id,
                "Controller trace id",
            )
            deadline_at = context.deadline_at
        else:
            raise ValueError("bridge context must be Controller Context or None")

        try:
            response = self._transport.exchange(
                {
                    "requestId": request_id,
                    "traceId": trace_id,
                    "operation": operation,
                    "arguments": arguments,
                },
                deadline_at,
            )
        except WriterBridgeTransportError as exc:
            raise WriterBackendActionFailure(
                outcome="unknown",
                code="RESPONSE_LOST",
                message="The Writer bridge did not return a reliable response",
                binding_disposition="unprovable",
            ) from exc

        try:
            response = _require_mapping(response, "Writer bridge response")
            if response.get("requestId") != request_id:
                raise TypeError(
                    "Writer bridge response requestId does not match"
                )

            outcome = response.get("outcome")
            if outcome == "succeeded":
                response = _require_exact_fields(
                    response,
                    {"requestId", "outcome", "data"},
                    "successful Writer bridge response",
                )
                return _require_mapping(
                    response["data"],
                    "successful Writer bridge data",
                )
            if outcome not in {"failed", "unknown"}:
                raise TypeError("Writer bridge response has invalid outcome")
            response = _require_exact_fields(
                response,
                {
                    "requestId",
                    "outcome",
                    "error",
                    "bindingDisposition",
                },
                "failed Writer bridge response",
            )
            error = _require_exact_fields(
                response["error"],
                {"code", "message"},
                "Writer bridge error",
            )
            raise WriterBackendActionFailure(
                outcome=outcome,
                code=_require_nonempty_string(
                    error["code"],
                    "Writer bridge error code",
                ),
                message=_require_nonempty_string(
                    error["message"],
                    "Writer bridge error message",
                ),
                binding_disposition=response["bindingDisposition"],
            )
        except WriterBackendActionFailure:
            raise
        except (TypeError, ValueError) as exc:
            raise WriterBackendActionFailure(
                outcome="unknown",
                code="RESPONSE_LOST",
                message="The Writer bridge returned an invalid response",
                binding_disposition="unprovable",
            ) from exc

    def close(self):
        return self._transport.close()
