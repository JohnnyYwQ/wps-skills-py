#!/usr/bin/env python3
"""
WPS 统一 Skill 调用入口（胶水层）

模型/用户只需调用本脚本即可操控 WPS Excel / PPT / Word，无需手动启动桥接服务：

    python scripts/call.py getContext '{}'                       # Excel
    python scripts/call.py createPresentation '{}'               # PPT
    python scripts/call.py setFont '{"font_name":"微软雅黑"}'      # Word
    python scripts/call.py save --app ppt '{}'                  # 通用（按 app 委派）

本脚本会：
  1. 检查桥接服务是否在运行（GET /health）
  2. 若未运行，自动在后台拉起 bridge/server.py（无需用户手动起服务）
  3. POST action 到 http://127.0.0.1:58891/dispatch 并返回 JSON 结果
"""

from dataclasses import dataclass
from contextlib import ExitStack, contextmanager
import errno
import sys
import os
import json
import re
import socket
import tempfile
import time
import subprocess
import uuid
import urllib.request
import urllib.error

BRIDGE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "bridge")
if BRIDGE_DIR not in sys.path:
    sys.path.insert(0, BRIDGE_DIR)

from action_trace import ActionTrace, server_log_path
from service_lifecycle import (
    ServiceStartResult,
    bridge_host,
    bridge_port,
    current_service_identity,
    inspect_health,
    utc_timestamp,
)


HOST = bridge_host()
PORT = bridge_port()
BASE = f"http://{HOST}:{PORT}"
_LOOPBACK_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
_DIAGNOSTIC_POPEN = subprocess.Popen


def _urlopen(request, timeout):
    """直连 loopback bridge，不继承 HTTP_PROXY/HTTPS_PROXY。"""
    return _LOOPBACK_OPENER.open(request, timeout=timeout)


@dataclass(frozen=True)
class HealthProbe:
    state: str
    health: object = None
    error: str = None


@dataclass(frozen=True)
class HealthObservation:
    probe: HealthProbe
    listener_before: object = None
    listener_after: object = None


def _as_health_probe(value):
    """兼容测试/旧内部调用传入的 dict/None。"""
    if isinstance(value, HealthProbe):
        return value
    if value is None:
        return HealthProbe("absent")
    return HealthProbe("responded", health=value)


def _process_created_at(pid):
    if not pid:
        return None
    try:
        if sys.platform == "win32":
            stdout = _run_diagnostic(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    f"(Get-CimInstance Win32_Process -Filter 'ProcessId = {int(pid)}').CreationDate",
                ],
            )
        else:
            stdout = _run_diagnostic(["ps", "-p", str(int(pid)), "-o", "lstart="])
        created_at = stdout.strip()
        return created_at or None
    except Exception:
        return None


def _listener_identity():
    """Best-effort identity of the process listening on the bridge port."""
    try:
        if sys.platform == "win32":
            stdout = _run_diagnostic(["netstat", "-ano", "-p", "tcp"])
            pattern = re.compile(
                rf"^\s*TCP\s+{re.escape(HOST)}:{PORT}\s+\S+\s+LISTENING\s+(\d+)\s*$",
                re.IGNORECASE | re.MULTILINE,
            )
            match = pattern.search(stdout)
            pid = int(match.group(1)) if match else None
        else:
            stdout = _run_diagnostic(
                ["lsof", "-nP", "-a", f"-iTCP:{PORT}", "-sTCP:LISTEN", "-Fp"]
            )
            match = re.search(r"^p(\d+)$", stdout, re.MULTILINE)
            pid = int(match.group(1)) if match else None
        return {"pid": pid, "createdAt": _process_created_at(pid)}
    except Exception as exc:
        return {
            "pid": None,
            "createdAt": None,
            "inspectionError": f"{type(exc).__name__}: {exc}",
        }


def _run_diagnostic(command, timeout=2):
    process = _DIAGNOSTIC_POPEN(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    try:
        stdout, _ = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        process.kill()
        stdout, _ = process.communicate()
    return stdout or ""


def _observe_health():
    listener_before = _listener_identity()
    probe = _as_health_probe(_health(with_state=True))
    listener_after = _listener_identity()
    return HealthObservation(probe, listener_before, listener_after)


def _is_connection_refused(exc):
    refused_codes = {errno.ECONNREFUSED, 10061}
    current = exc
    for _ in range(3):
        if isinstance(current, ConnectionRefusedError):
            return True
        if getattr(current, "errno", None) in refused_codes:
            return True
        reason = getattr(current, "reason", None)
        if reason is None or reason is current:
            break
        current = reason
    return False


def _project_root():
    """scripts/ 的上级目录 = 技能根目录"""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _server_path():
    return os.path.join(_project_root(), "bridge", "server.py")


@contextmanager
def _startup_lock():
    """Serialize bridge discovery and startup across independent callers."""
    safe_host = HOST.replace(":", "_").replace(".", "_")
    lock_path = os.path.join(
        tempfile.gettempdir(),
        f"wps-skills-bridge-{safe_host}-{PORT}.lock",
    )
    lock_file = open(lock_path, "a+b")
    locked = False
    try:
        if sys.platform == "win32":
            import msvcrt

            lock_file.seek(0, os.SEEK_END)
            if lock_file.tell() == 0:
                lock_file.write(b"\0")
                lock_file.flush()
            lock_file.seek(0)
            while True:
                try:
                    msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
                    break
                except OSError as exc:
                    if exc.errno not in (errno.EACCES, errno.EAGAIN, errno.EDEADLK):
                        raise
                    time.sleep(0.05)
        else:
            import fcntl

            fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX)
        locked = True
        yield
    finally:
        try:
            if locked:
                lock_file.seek(0)
                if sys.platform == "win32":
                    import msvcrt

                    msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl

                    fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)
        finally:
            lock_file.close()


def _health(with_state=False):
    try:
        with _urlopen(f"{BASE}/health", timeout=2) as resp:
            raw = resp.read().decode("utf-8")
        try:
            probe = HealthProbe("responded", health=json.loads(raw))
        except Exception as exc:
            probe = HealthProbe(
                "unhealthy",
                error=f"bridge /health 返回非 JSON: {type(exc).__name__}: {exc}",
            )
    except urllib.error.HTTPError as exc:
        try:
            health = json.loads(exc.read().decode("utf-8"))
            probe = HealthProbe("responded", health=health, error=f"HTTP {exc.code}")
        except Exception:
            probe = HealthProbe("unhealthy", error=f"bridge /health 返回 HTTP {exc.code}")
        finally:
            exc.close()
    except Exception as exc:
        if _is_connection_refused(exc):
            probe = HealthProbe("absent", error=str(exc))
        else:
            probe = HealthProbe(
                "unresponsive",
                error=f"{type(exc).__name__}: {exc}",
            )
    return probe if with_state else probe.health


def _is_healthy(h, expected=None):
    """只有 checkout 与运行时代码都匹配的 bridge 才可复用。"""
    return inspect_health(h, expected=expected).reusable


def _terminate_started_process(process, wait_seconds=3):
    """回收本次调用亲自启动但未就绪的进程；不触碰已有陌生实例。"""
    if process is None:
        return
    try:
        if process.poll() is not None:
            return
    except Exception:
        # poll 本身失败时无法证明子进程已退出，继续做有界回收。
        pass
    try:
        process.terminate()
        process.wait(timeout=wait_seconds)
        return
    except subprocess.TimeoutExpired:
        pass
    except Exception:
        pass
    try:
        process.kill()
        process.wait(timeout=wait_seconds)
    except Exception:
        pass


def _observable_identity_error(health):
    if not isinstance(health, dict):
        return "bridge health 缺少可观察的启动身份"
    positive_pid_fields = ("pid", "serverPid", "launchedByPid")
    for field in positive_pid_fields:
        value = health.get(field)
        if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
            return f"bridge health 缺少有效的 {field}"
    if health["pid"] != health["serverPid"]:
        return "bridge health 的 pid 与 serverPid 不一致"
    for field in ("launchId", "serverPath", "launchStartedAt"):
        if not isinstance(health.get(field), str) or not health[field].strip():
            return f"bridge health 缺少有效的 {field}"
    expected_server = os.path.normcase(os.path.realpath(_server_path()))
    actual_server = os.path.normcase(os.path.realpath(health["serverPath"]))
    if actual_server != expected_server:
        return f"bridge health 的 serverPath 不属于当前 checkout: {health['serverPath']}"
    return None


def _listener_ownership_error(observation, health):
    target_pid = health.get("serverPid") if isinstance(health, dict) else None
    before = observation.listener_before if isinstance(observation, HealthObservation) else None
    after = observation.listener_after if isinstance(observation, HealthObservation) else None
    if not isinstance(after, dict) or after.get("pid") != target_pid:
        actual_pid = after.get("pid") if isinstance(after, dict) else None
        return f"health 的 serverPid={target_pid} 与 probe 后监听 PID={actual_pid} 不一致"
    if not after.get("createdAt"):
        return f"无法确认监听 PID {target_pid} 的进程创建时间"
    if isinstance(before, dict) and before.get("pid") is not None:
        if before.get("pid") != target_pid:
            return (
                f"probe 期间监听 PID 从 {before.get('pid')} 变为 {target_pid}，"
                "无法确认端口所有权"
            )
        if before.get("createdAt") != after.get("createdAt"):
            return f"监听 PID {target_pid} 的进程创建时间在 probe 期间发生变化"
    return None


def _classify_existing_probe(observation, expected, trace=None, started=None):
    """Return a terminal result, or None when the port is confirmed absent."""
    if isinstance(observation, HealthObservation):
        probe = observation.probe
        listener_before = observation.listener_before
        listener_after = observation.listener_after
    else:
        probe = _as_health_probe(observation)
        listener_before = None
        listener_after = None
    health = probe.health
    listener_was_present = any(
        isinstance(snapshot, dict) and snapshot.get("pid")
        for snapshot in (listener_before, listener_after)
    )
    if probe.state == "absent" and listener_was_present:
        error = (
            "health probe 未确认 bridge，但探测前后发现端口监听者；"
            "为避免双开，拒绝启动第二实例"
        )
        if trace:
            trace.event(
                "bridge.health.unavailable",
                status="error",
                state="unknown_owner",
                error=error,
                listenerBefore=listener_before,
                listenerAfter=listener_after,
            )
        return ServiceStartResult(
            False,
            code="BRIDGE_UNAVAILABLE",
            error=error,
            disposition="unknown_owner",
            listener_before=listener_before,
            listener_after=listener_after,
        )
    if probe.state in ("unresponsive", "unhealthy"):
        error = (
            "bridge 端口存在但健康检查无响应，可能正在执行 Action 或已卡住；"
            "为避免双开，拒绝启动第二实例"
        )
        if probe.error:
            error = f"{error}: {probe.error}"
        if trace:
            trace.event(
                "bridge.health.unavailable",
                status="error",
                state=probe.state,
                error=error,
                listenerBefore=listener_before,
                listenerAfter=listener_after,
            )
        return ServiceStartResult(
            False,
            code="BRIDGE_UNAVAILABLE",
            error=error,
            disposition="unknown_owner",
            listener_before=listener_before,
            listener_after=listener_after,
        )
    inspection = inspect_health(health, expected=expected)
    if inspection.reusable:
        identity_error = _observable_identity_error(health) or _listener_ownership_error(
            observation,
            health,
        )
        if identity_error:
            if trace:
                trace.event(
                    "bridge.instance.rejected",
                    status="error",
                    state="unknown_owner",
                    error=identity_error,
                    listenerBefore=listener_before,
                    listenerAfter=listener_after,
                )
            return ServiceStartResult(
                False,
                code="BRIDGE_INSTANCE_MISMATCH",
                error=identity_error,
                health=health,
                disposition="unknown_owner",
                listener_before=listener_before,
                listener_after=listener_after,
            )
        if trace:
            trace.event(
                "bridge.health.checked",
                status="healthy",
                instanceId=health.get("instanceId"),
                elapsedMs=(
                    round((time.perf_counter() - started) * 1000, 2)
                    if started is not None
                    else None
                ),
            )
        return ServiceStartResult(True, health=health, disposition="reused")
    if probe.state == "responded":
        inspection_error = inspection.error
        if health is None:
            inspection_error = "端口上的 /health 返回 JSON null，拒绝视为空闲端口"
        if trace:
            trace.event(
                "bridge.instance.rejected",
                status="error",
                state=inspection.state,
                error=inspection_error,
                expectedProjectRoot=expected.get("projectRoot"),
                actualProjectRoot=(health.get("projectRoot") if isinstance(health, dict) else None),
                actualInstanceId=(health.get("instanceId") if isinstance(health, dict) else None),
            )
        return ServiceStartResult(
            False,
            code="BRIDGE_INSTANCE_MISMATCH",
            error=inspection_error,
            health=health if isinstance(health, dict) else None,
            disposition=("foreign" if inspection.state == "foreign" else "unknown_owner"),
        )
    return None


def _ensure_server(trace=None):
    """确保当前 checkout 的 bridge 在运行；绝不复用身份不明的服务。"""
    started = time.perf_counter()
    expected = current_service_identity()
    initial_observation = _observe_health()
    result = _classify_existing_probe(
        initial_observation,
        expected,
        trace=trace,
        started=started,
    )
    if result is not None:
        return result

    startup_stack = ExitStack()
    try:
        startup_stack.enter_context(_startup_lock())
    except Exception as exc:
        error = f"无法获得 bridge 跨进程启动锁: {type(exc).__name__}: {exc}"
        if trace:
            trace.event(
                "bridge.start.lock_failed",
                status="error",
                error=error,
                listenerBefore=initial_observation.listener_before,
                listenerAfter=initial_observation.listener_after,
            )
        return ServiceStartResult(
            False,
            code="BRIDGE_START_LOCK_FAILED",
            error=error,
            disposition="unknown_owner",
            listener_before=initial_observation.listener_before,
            listener_after=initial_observation.listener_after,
        )
    with startup_stack:
        # 获得启动权后重新探测，关闭 probe 与 spawn 之间的 TOCTOU 窗口。
        result = _classify_existing_probe(
            _observe_health(),
            expected,
            trace=trace,
            started=started,
        )
        if result is not None:
            return result

        return _start_server_locked(expected, trace=trace, started=started)


def _start_server_locked(expected, trace=None, started=None):
    """Start one bridge while the caller owns the cross-process startup lock."""
    if started is None:
        started = time.perf_counter()

    py = sys.executable
    server = _server_path()
    if not os.path.exists(server):
        error = f"找不到桥接服务: {server}"
        if trace:
            trace.event("bridge.start.failed", status="error", error=error)
        return ServiceStartResult(False, code="BRIDGE_START_FAILED", error=error)

    launch_id = uuid.uuid4().hex
    launch_started_at = utc_timestamp()
    launch_environment = os.environ.copy()
    launch_environment.update(
        {
            "WPS_BRIDGE_LAUNCH_ID": launch_id,
            "WPS_BRIDGE_LAUNCHED_BY_PID": str(os.getpid()),
            "WPS_BRIDGE_SERVER_PATH": server,
            "WPS_BRIDGE_PROJECT_ROOT": _project_root(),
            "WPS_BRIDGE_LAUNCH_STARTED_AT": launch_started_at,
        }
    )

    log_path, log_warning = server_log_path()
    if trace:
        trace.event(
            "bridge.start.requested",
            launchId=launch_id,
            callerPid=os.getpid(),
            serverPath=server,
            projectRoot=_project_root(),
            launchStartedAt=launch_started_at,
            serverLog=str(log_path) if log_path else None,
        )
        if log_warning:
            trace.event("bridge.log.warning", status="warning", warning=log_warning)

    log_stream = None
    process = None
    try:
        if log_path is not None:
            log_stream = log_path.open("a", encoding="utf-8")
        stdout_target = log_stream if log_stream is not None else subprocess.DEVNULL
        stderr_target = subprocess.STDOUT if log_stream is not None else subprocess.DEVNULL
        if os.name == "nt":
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            process = subprocess.Popen(
                [py, server],
                stdout=stdout_target,
                stderr=stderr_target,
                creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
                startupinfo=si,
                env=launch_environment,
            )
        else:
            process = subprocess.Popen(
                [py, server],
                stdout=stdout_target,
                stderr=stderr_target,
                start_new_session=True,
                env=launch_environment,
            )
        if trace:
            trace.event(
                "bridge.start.spawned",
                launchId=launch_id,
                callerPid=os.getpid(),
                processPid=getattr(process, "pid", None),
                serverPath=server,
                projectRoot=_project_root(),
                launchStartedAt=launch_started_at,
            )
    except Exception as e:
        error = f"启动桥接服务失败: {e}"
        if trace:
            trace.event("bridge.start.failed", status="error", error=error)
        return ServiceStartResult(False, code="BRIDGE_START_FAILED", error=error)
    finally:
        if log_stream is not None:
            log_stream.close()

    # 轮询等待服务就绪（最多 ~10s）
    last_observation = None
    for _ in range(40):
        time.sleep(0.25)
        observation = _observe_health()
        last_observation = observation
        probe = observation.probe
        health = probe.health
        inspection = inspect_health(health, expected=expected)
        if inspection.reusable:
            process_pid = getattr(process, "pid", None)
            response_pid = health.get("serverPid", health.get("pid"))
            identity_error = _observable_identity_error(health) or _listener_ownership_error(
                observation,
                health,
            )
            if identity_error:
                _terminate_started_process(process)
                if trace:
                    trace.event(
                        "bridge.start.conflicted",
                        status="error",
                        state="unknown_owner",
                        error=identity_error,
                        launchId=launch_id,
                        processPid=process_pid,
                        listenerBefore=observation.listener_before,
                        listenerAfter=observation.listener_after,
                    )
                return ServiceStartResult(
                    False,
                    code="BRIDGE_INSTANCE_MISMATCH",
                    error=identity_error,
                    health=health,
                    disposition="unknown_owner",
                    listener_before=observation.listener_before,
                    listener_after=observation.listener_after,
                )
            is_started_process = bool(
                process_pid
                and health.get("launchId") == launch_id
                and response_pid == process_pid
                and health.get("pid") == process_pid
                and health.get("launchedByPid") == os.getpid()
                and health.get("launchStartedAt") == launch_started_at
            )
            if not is_started_process:
                _terminate_started_process(process)
                identity_conflicts_with_child = bool(
                    response_pid == process_pid or health.get("launchId") == launch_id
                )
                if trace:
                    trace.event(
                        (
                            "bridge.start.conflicted"
                            if identity_conflicts_with_child
                            else "bridge.start.superseded"
                        ),
                        status=("error" if identity_conflicts_with_child else "healthy"),
                        launchId=launch_id,
                        processPid=process_pid,
                        responseLaunchId=health.get("launchId"),
                        responsePid=response_pid,
                        instanceId=health.get("instanceId"),
                    )
                if identity_conflicts_with_child:
                    error = "READY bridge 的 PID/launchId 与本次启动记录不一致"
                    return ServiceStartResult(
                        False,
                        code="BRIDGE_INSTANCE_MISMATCH",
                        error=error,
                        health=health,
                        disposition="unknown_owner",
                        listener_before=observation.listener_before,
                        listener_after=observation.listener_after,
                    )
                return ServiceStartResult(
                    True,
                    health=health,
                    started=False,
                    disposition="reused",
                )
            if trace:
                trace.event(
                    "bridge.start.completed",
                    status="healthy",
                    launchId=launch_id,
                    processPid=process_pid,
                    instanceId=health.get("instanceId"),
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )
            return ServiceStartResult(
                True,
                health=health,
                started=True,
                disposition="self_started",
            )
        if probe.state == "responded":
            inspection_error = inspection.error
            if health is None:
                inspection_error = "端口上的 /health 返回 JSON null，拒绝视为空闲端口"
            _terminate_started_process(process)
            if trace:
                trace.event(
                    "bridge.start.conflicted",
                    status="error",
                    state=inspection.state,
                    error=inspection_error,
                    actualInstanceId=(health.get("instanceId") if isinstance(health, dict) else None),
                )
            return ServiceStartResult(
                False,
                code="BRIDGE_INSTANCE_MISMATCH",
                error=inspection_error,
                health=health if isinstance(health, dict) else None,
                disposition=("foreign" if inspection.state == "foreign" else "unknown_owner"),
                listener_before=observation.listener_before,
                listener_after=observation.listener_after,
            )
        try:
            process_exit_code = process.poll() if process is not None else None
        except Exception:
            process_exit_code = None
        if process_exit_code is not None:
            error = f"桥接服务启动后提前退出（exit={process_exit_code}），请检查 server 日志"
            if trace:
                trace.event("bridge.start.failed", status="error", error=error)
            return ServiceStartResult(
                False,
                code="BRIDGE_START_EXITED",
                error=error,
                disposition="unknown_owner",
                listener_before=observation.listener_before,
                listener_after=observation.listener_after,
            )

    _terminate_started_process(process)
    error = "桥接服务启动超时，已回收本次启动的后台进程"
    if trace:
        trace.event(
            "bridge.start.failed",
            status="timeout",
            error=error,
            elapsedMs=round((time.perf_counter() - started) * 1000, 2),
            serverLog=str(log_path) if log_path else None,
            launchId=launch_id,
            processPid=getattr(process, "pid", None),
            listenerBefore=(last_observation.listener_before if last_observation else None),
            listenerAfter=(last_observation.listener_after if last_observation else None),
        )
    return ServiceStartResult(
        False,
        code="BRIDGE_START_TIMEOUT",
        error=error,
        disposition="unknown_owner",
        listener_before=(last_observation.listener_before if last_observation else None),
        listener_after=(last_observation.listener_after if last_observation else None),
    )


def _post(action, params, app=None, trace=None, service_health=None):
    trace = trace or ActionTrace.start(component="call")
    service_health = service_health or _health()
    inspection = inspect_health(service_health)
    if not inspection.reusable:
        raise RuntimeError(inspection.error or "bridge 实例身份校验失败")
    payload = {"traceId": trace.trace_id, "action": action, "params": params}
    if app:
        payload["app"] = app
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        f"{BASE}/dispatch",
        data=body,
        headers={
            "Content-Type": "application/json",
            "X-WPS-Trace-Id": trace.trace_id,
            "X-WPS-Bridge-Project-Id": service_health["projectId"],
            "X-WPS-Bridge-Instance-Id": service_health["instanceId"],
        },
    )
    trace.event(
        "http.request.sent",
        method="POST",
        path="/dispatch",
        action=action,
        requestedApp=app,
        bodyBytes=len(body),
        **trace.debug_fields(params=params),
    )
    started = time.perf_counter()
    try:
        with _urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            http_status = getattr(resp, "status", 200)
    except urllib.error.HTTPError as exc:
        http_status = exc.code
        try:
            result = json.loads(exc.read().decode("utf-8"))
        except Exception as parse_exc:
            trace.event(
                "http.request.failed",
                status="error",
                action=action,
                error=f"HTTP {exc.code} 返回非 JSON: {parse_exc}",
                elapsedMs=round((time.perf_counter() - started) * 1000, 2),
            )
            raise RuntimeError(f"bridge HTTP {exc.code} 返回非 JSON") from parse_exc
        finally:
            exc.close()
    except Exception as exc:
        trace.event(
            "http.request.failed",
            status="error",
            action=action,
            error=f"{type(exc).__name__}: {exc}",
            elapsedMs=round((time.perf_counter() - started) * 1000, 2),
        )
        raise

    if not isinstance(result, dict):
        result = {"success": False, "error": "桥接服务返回的不是 JSON object"}
    trace.event(
        "http.response.received",
        status="success" if result.get("success") else "error",
        httpStatus=http_status,
        action=action,
        code=result.get("code"),
        error=result.get("error"),
        elapsedMs=round((time.perf_counter() - started) * 1000, 2),
        **trace.debug_fields(response=result),
    )
    return trace.decorate(result)


def _usage_error(msg):
    print(json.dumps({"success": False, "error": msg}, ensure_ascii=False))
    sys.exit(1)


def _load_params_from_args(args):
    """解析调用参数，支持三种等价写法：

    1) 旧式（bash 友好）:  call.py <action> '<json>'
    2) 文件（PowerShell 5.1 友好，彻底避开命令行引号转义）:
                         call.py <action> --params-file <path.json>
    3) 管道:             echo '<json>' | call.py <action> [--stdin]
                         call.py <action> --stdin < path.json

    可选路由参数:      call.py <action> --app <excel|ppt|word> ...

    返回 (action, params, app)。"""
    action = None
    params = None
    app = None
    use_stdin = False
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--stdin":
            use_stdin = True
            i += 1
            continue
        if a == "--app":
            i += 1
            if i >= len(args):
                _usage_error("--app 缺少应用参数（excel/ppt/word）")
            new_app = args[i].strip().lower()
            if app is not None and app != new_app:
                _usage_error(f"重复的 --app 参数不一致: {app} / {new_app}")
            app = new_app
            i += 1
            continue
        if a == "--params-file":
            i += 1
            if i >= len(args):
                _usage_error("--params-file 缺少文件路径参数")
            path = args[i]
            try:
                with open(path, "r", encoding="utf-8") as f:
                    params = json.load(f)
            except FileNotFoundError:
                _usage_error(f"--params-file 指定的文件不存在: {path}")
            except json.JSONDecodeError as e:
                _usage_error(f"--params-file 内容不是合法 JSON: {e}")
            except Exception as e:
                _usage_error(f"读取 --params-file 失败: {e}")
            i += 1
            continue
        # 位置参数：第一个是 action，第二个（若仍是位置参数）视为 JSON 字符串
        if action is None:
            action = a
        elif params is None:
            try:
                params = json.loads(a)
            except json.JSONDecodeError as e:
                _usage_error(f"params 不是合法 JSON（PowerShell 下建议改用 --params-file/--stdin）: {e}")
        i += 1

    # 未显式获得 params 时，按需回退到 stdin（管道场景）
    if params is None and (use_stdin or (not sys.stdin.isatty())):
        try:
            raw = sys.stdin.read()
            if raw.strip():
                params = json.loads(raw)
        except Exception:
            pass

    if action is None:
        _usage_error("用法: call.py <action> [--app excel|ppt|word] ['<json>' | --params-file <path> | --stdin]")
    if params is None:
        params = {}
    return action, params, app


def _bridge_ensure_summary(service):
    health = service.health if isinstance(service.health, dict) else {}
    return {
        "disposition": service.disposition,
        "instanceId": health.get("instanceId"),
        "launchId": health.get("launchId"),
        "serverPid": health.get("serverPid"),
        "listenerBefore": service.listener_before,
        "listenerAfter": service.listener_after,
    }


def main():
    action, params, app = _load_params_from_args(sys.argv[1:])
    trace = ActionTrace.start(component="call")
    action_started = time.perf_counter()
    trace.event(
        "action.started",
        action=action,
        requestedApp=app or (params.get("app") if isinstance(params, dict) else None),
        **trace.debug_fields(params=params),
    )

    service = _ensure_server(trace=trace)
    bridge_ensure = _bridge_ensure_summary(service)
    if not service:
        result = {
            "success": False,
            "code": service.code,
            "error": service.error or "桥接服务启动失败，请检查 traceLog 与 server 日志",
            "bridgeEnsure": bridge_ensure,
        }
        trace.event(
            "action.completed",
            status="error",
            action=action,
            elapsedMs=round((time.perf_counter() - action_started) * 1000, 2),
        )
        print(json.dumps(trace.decorate(result), ensure_ascii=False))
        sys.exit(1)

    try:
        result = _post(
            action,
            params,
            app=app,
            trace=trace,
            service_health=service.health,
        )
        result["bridgeEnsure"] = bridge_ensure
    except urllib.error.URLError as e:
        result = {
            "success": False,
            "error": f"调用失败（服务无响应）: {e}",
            "bridgeEnsure": bridge_ensure,
        }
        trace.event(
            "action.completed",
            status="error",
            action=action,
            elapsedMs=round((time.perf_counter() - action_started) * 1000, 2),
        )
        print(json.dumps(trace.decorate(result), ensure_ascii=False))
        sys.exit(1)
    except Exception as e:
        result = {
            "success": False,
            "error": f"调用失败: {e}",
            "bridgeEnsure": bridge_ensure,
        }
        trace.event(
            "action.completed",
            status="error",
            action=action,
            elapsedMs=round((time.perf_counter() - action_started) * 1000, 2),
        )
        print(json.dumps(trace.decorate(result), ensure_ascii=False))
        sys.exit(1)

    trace.event(
        "action.completed",
        status="success" if result.get("success") else "error",
        action=action,
        code=result.get("code"),
        error=result.get("error"),
        elapsedMs=round((time.perf_counter() - action_started) * 1000, 2),
    )
    result = trace.decorate(result)
    print(json.dumps(result, ensure_ascii=False))
    sys.exit(0 if result.get("success") else 1)


if __name__ == "__main__":
    main()
