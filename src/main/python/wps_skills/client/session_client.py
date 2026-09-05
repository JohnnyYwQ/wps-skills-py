"""Own one Host channel and decide each Action only after its complete response."""

from collections import deque
import json
import math
import os
import queue
import subprocess
import threading
import time


class SessionClientError(RuntimeError):
    """Channel/startup/cleanup failure; no Action Response is invented."""

    def __init__(self, message, *, may_have_effect=False):
        super().__init__(message)
        self.may_have_effect = may_have_effect


class ActionFailed(RuntimeError):
    """An authoritative failed or unknown Action Response, preserved intact."""

    def __init__(self, response):
        self.response = response
        error = response["error"]
        super().__init__(f'{response["outcome"]}: {error["code"]}: {error["message"]}')


def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("duplicate JSON key")
        result[key] = value
    return result


def _nonfinite(value):
    raise ValueError("non-finite JSON number")


class SessionClient:
    """Single-caller client. Failure never retries, rebinds, or starts a new Host.

    call(address, params) returns the canonical successful Action Response or
    raises ActionFailed with the failed/unknown response. session_outcome and
    cleanup_error describe cleanup separately, including after an Action error.
    """

    def __init__(self, command, *, application, cwd=None, env=None, timeout=60):
        if not command or isinstance(command, str):
            raise ValueError("command must be a nonempty argument sequence")
        if application not in {"word", "excel", "ppt"}:
            raise ValueError("unknown application")
        if not isinstance(timeout, (int, float)) or not math.isfinite(timeout) or timeout <= 0:
            raise ValueError("timeout must be a positive finite number")
        self._command = list(command)
        self._application = application
        self._cwd = cwd
        self._env = env
        self._timeout = timeout
        self._process = None
        self._records = queue.Queue()
        self._stderr = deque(maxlen=64)
        self._threads = []
        self._lock = threading.Lock()
        self._started = False
        self._terminal = False
        self.ready = None
        self.last_response = None
        self.session_outcome = None
        self.cleanup_error = None

    @property
    def can_execute(self):
        return self.ready is not None and not self._terminal

    @property
    def stderr(self):
        return "".join(self._stderr)

    def _read_stdout(self):
        try:
            for line in self._process.stdout:
                self._records.put(line)
        finally:
            self._records.put(None)

    def _read_stderr(self):
        for line in self._process.stderr:
            self._stderr.append(line)

    def _receive(self, deadline):
        try:
            line = self._records.get(timeout=max(0, deadline - time.monotonic()))
        except queue.Empty as exc:
            raise SessionClientError("Timed out waiting for the Session Host; do not replay the Action") from exc
        if line is None:
            raise SessionClientError("Session Host channel ended; inspect stderr and traces before continuing")
        try:
            record = json.loads(line, object_pairs_hook=_unique_object, parse_constant=_nonfinite)
            if not isinstance(record, dict):
                raise ValueError("record is not an object")
            # Reject floating point overflow too, without reinterpreting data.
            json.dumps(record, allow_nan=False)
            return record
        except (ValueError, TypeError) as exc:
            raise SessionClientError("Invalid Session Protocol record") from exc

    def _send(self, encoded):
        try:
            self._process.stdin.write(encoded + "\n")
            self._process.stdin.flush()
        except (OSError, ValueError) as exc:
            raise SessionClientError("Session Host input channel failed") from exc

    def start(self):
        if self._started:
            raise SessionClientError("A Session Client can start only once")
        self._started = True
        try:
            self._process = subprocess.Popen(
                self._command, cwd=self._cwd, env=self._env,
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, encoding="utf-8", bufsize=1,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
            )
            for target in (self._read_stdout, self._read_stderr):
                thread = threading.Thread(target=target, daemon=True)
                thread.start()
                self._threads.append(thread)
            ready = self._receive(time.monotonic() + self._timeout)
            if (ready.get("type") != "session.ready"
                    or ready.get("protocolVersion") != 1
                    or ready.get("app") != self._application
                    or not isinstance(ready.get("sessionId"), str)
                    or not ready["sessionId"]):
                raise SessionClientError("Host did not establish the requested Protocol v1 Session")
            self.ready = ready
            return self
        except BaseException:
            self._abort()
            raise

    def _accept_closed(self, record):
        self._terminal = True
        outcome = record.get("session")
        if (record.get("type") != "session.closed" or not isinstance(outcome, dict)
                or outcome.get("sessionId") != self.ready["sessionId"]
                or outcome.get("app") != self._application
                or outcome.get("outcome") not in {"succeeded", "failed"}):
            raise SessionClientError("Invalid session.closed record")
        self.session_outcome = outcome

    def _finish_process(self, deadline):
        self._terminal = True
        try:
            self._process.stdin.close()
        except OSError:
            pass
        try:
            self._process.wait(timeout=max(0.01, deadline - time.monotonic()))
        except subprocess.TimeoutExpired as exc:
            raise SessionClientError("Session Host did not exit after closure") from exc
        for thread in self._threads:
            thread.join(timeout=1)
        self._process.stdout.close()
        self._process.stderr.close()
        if self.session_outcome["outcome"] != "succeeded":
            raise SessionClientError("Session cleanup failed; inspect session_outcome")
        # A Host failure exit bit accompanies legitimate terminal Action errors.
        if self._process.returncode and self.session_outcome.get("reason") != "host_failure":
            raise SessionClientError("Session Host exited with an unexpected failure code")

    def _abort(self):
        self._terminal = True
        if self._process is None:
            return
        try:
            self._process.stdin.close()
        except (OSError, ValueError):
            pass
        try:
            self._process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            # Only the owned Host is terminated. Never close a WPS application.
            self._process.terminate()
            try:
                self._process.wait(timeout=2)
            except subprocess.TimeoutExpired:
                self._process.kill()
                self._process.wait(timeout=2)
        for thread in self._threads:
            thread.join(timeout=1)
        for stream in (self._process.stdout, self._process.stderr):
            stream.close()

    def call(self, address, params):
        """Submit exactly one request; unsuccessful responses stop ordinary scripts."""
        if (not isinstance(address, dict) or set(address) != {"app", "action"}
                or address["app"] != self._application
                or not isinstance(address["action"], str) or not address["action"]
                or not isinstance(params, dict)):
            raise ValueError("call requires an explicit matching app/action address and object params")
        encoded = json.dumps({"address": address, "params": params}, ensure_ascii=True, allow_nan=False)
        if not self._lock.acquire(blocking=False):
            raise SessionClientError("Another operation is already in flight")
        dispatched = False
        response = None
        try:
            if not self.can_execute:
                raise SessionClientError("Session is not available; it cannot be restarted or rebound")
            deadline = time.monotonic() + self._timeout
            # Detect idle closure already received before writing a new Action.
            if not self._records.empty():
                record = self._receive(deadline)
                self._accept_closed(record)
                self._finish_process(deadline)
                raise SessionClientError("Session closed while idle; no Action was submitted")
            dispatched = True
            self._send(encoded)
            record = self._receive(deadline)
            closing = record.get("type") == "session.closing"
            if closing:
                self._terminal = True
                if record.get("sessionId") != self.ready["sessionId"]:
                    raise SessionClientError("Mismatched session.closing identity")
                record = self._receive(deadline)
            if record.get("type") == "session.closed":
                self._accept_closed(record)
                self._finish_process(deadline)
                raise SessionClientError("Session closed without an Action Response")
            if (record.get("address") != address
                    or record.get("sessionId") != self.ready["sessionId"]
                    or record.get("outcome") not in {"succeeded", "failed", "unknown"}
                    or not isinstance(record.get("traceId"), str)):
                raise SessionClientError("Invalid or mismatched Action Response")
            if record["outcome"] == "succeeded":
                if not isinstance(record.get("data"), dict):
                    raise SessionClientError("Successful Action Response has no data")
            elif (not isinstance(record.get("error"), dict)
                    or not isinstance(record["error"].get("code"), str)
                    or not isinstance(record["error"].get("message"), str)):
                raise SessionClientError("Unsuccessful Action Response has no error")
            response = self.last_response = record
            if closing:
                try:
                    self._accept_closed(self._receive(deadline))
                    self._finish_process(deadline)
                except SessionClientError as exc:
                    self.cleanup_error = exc
                    self._abort()
            if response["outcome"] != "succeeded":
                raise ActionFailed(response)
            if self.cleanup_error is not None:
                raise self.cleanup_error
            return response
        except SessionClientError as exc:
            exc.may_have_effect = dispatched and response is None
            self._abort()
            raise
        except ActionFailed:
            raise
        except BaseException:
            # An interrupted caller must not leave a live Action channel behind.
            self._abort()
            raise
        finally:
            self._lock.release()

    def close(self):
        """Release automation resources; never save or close the WPS document."""
        if not self._lock.acquire(blocking=False):
            raise SessionClientError("Cannot close while an Action is in flight")
        try:
            if self._terminal or not self._started:
                if self.cleanup_error is not None:
                    raise self.cleanup_error
                return self.session_outcome
            deadline = time.monotonic() + self._timeout
            if self._records.empty():
                self._send('{"control":"close"}')
            self._accept_closed(self._receive(deadline))
            self._finish_process(deadline)
            return self.session_outcome
        except SessionClientError as exc:
            self.cleanup_error = exc
            self._abort()
            raise
        finally:
            self._lock.release()

    def __enter__(self):
        return self.start()

    def __exit__(self, exc_type, exc, tb):
        try:
            self.close()
        except SessionClientError:
            if exc_type is None:
                raise
        return False
