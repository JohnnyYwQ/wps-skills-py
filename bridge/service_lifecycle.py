#!/usr/bin/env python3
"""Bridge 实例身份与生命周期。

调用端和服务端共享同一份身份算法，只有 checkout 与运行时代码都匹配的
bridge 才能复用。BridgeLifecycle 封装显式停止与空闲回收，不持有 WPS 任务
语义；编排者仍负责在整个任务结束后调用 service.py stop。
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import hashlib
import os
from pathlib import Path
import subprocess
import threading
import time
from typing import Callable, Optional, Union
import uuid


SERVICE_NAME = "wps-skills-bridge"
PROTOCOL_VERSION = 1
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 58891

# 任务结束时应由编排者显式 stop；该超时只处理异常中断或遗漏清理。
# 设为 0 可关闭自动回收。
DEFAULT_IDLE_TIMEOUT_SECONDS = 15 * 60

# PowerShell line bridge 正常退出与强制回收的最长等待时间。正常关闭先让脚本
# 处理 EXIT，以便释放 COM；只有脚本卡死时才升级为 terminate/kill。
GRACEFUL_PROCESS_TIMEOUT_SECONDS = 2.0
FORCED_PROCESS_TIMEOUT_SECONDS = 1.0


def stop_line_process(
    process,
    graceful_timeout: float = GRACEFUL_PROCESS_TIMEOUT_SECONDS,
    forced_timeout: float = FORCED_PROCESS_TIMEOUT_SECONDS,
) -> None:
    """停止持久 line-RPC 子进程，优先走协议内的正常退出。"""
    if process is None:
        return

    try:
        running = process.poll() is None
    except Exception:
        running = True
    if not running:
        return

    try:
        if process.stdin:
            process.stdin.write("EXIT\n")
            process.stdin.flush()
    except Exception:
        # stdin 损坏不代表进程已退出，继续进入有界等待与回收流程。
        pass

    try:
        process.wait(timeout=graceful_timeout)
        return
    except subprocess.TimeoutExpired:
        pass
    except Exception:
        pass

    try:
        process.terminate()
        process.wait(timeout=forced_timeout)
        return
    except subprocess.TimeoutExpired:
        pass
    except Exception:
        pass

    try:
        process.kill()
        process.wait(timeout=forced_timeout)
    except Exception:
        # close() 是清理路径；目标进程可能已被其他线程或系统回收。
        pass


def _utc_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def bridge_host() -> str:
    # Bridge 只允许 loopback，避免把无认证的 WPS 控制端口暴露到网络。
    return DEFAULT_HOST


def bridge_port() -> int:
    raw = os.environ.get("WPS_BRIDGE_PORT", str(DEFAULT_PORT)).strip()
    try:
        port = int(raw)
    except (TypeError, ValueError):
        return DEFAULT_PORT
    return port if 1 <= port <= 65535 else DEFAULT_PORT


def configured_idle_timeout() -> float:
    raw = os.environ.get("WPS_BRIDGE_IDLE_SECONDS", str(DEFAULT_IDLE_TIMEOUT_SECONDS)).strip()
    try:
        return max(0.0, float(raw))
    except (TypeError, ValueError):
        return float(DEFAULT_IDLE_TIMEOUT_SECONDS)


def _runtime_files(root: Path):
    bridge_dir = root / "bridge"
    scripts_dir = root / "scripts"
    candidates = []
    if bridge_dir.is_dir():
        candidates.extend(path for path in bridge_dir.glob("*.py") if not path.name.startswith("test_"))
    vendor_dir = root / "vendor"
    for package in ("openpyxl", "et_xmlfile"):
        package_dir = vendor_dir / package
        if package_dir.is_dir():
            candidates.extend(package_dir.rglob("*.py"))
    for name in ("call.py", "service.py", "start.py"):
        candidates.append(scripts_dir / name)
    return sorted(
        (path for path in candidates if path.is_file()),
        key=lambda path: path.relative_to(root).as_posix(),
    )


def code_fingerprint(root: Optional[Union[str, os.PathLike]] = None) -> str:
    root_path = Path(root).resolve() if root is not None else project_root()
    digest = hashlib.sha256()
    for path in _runtime_files(root_path):
        relative = path.relative_to(root_path).as_posix().encode("utf-8")
        digest.update(len(relative).to_bytes(4, "big"))
        digest.update(relative)
        content = path.read_bytes()
        digest.update(len(content).to_bytes(8, "big"))
        digest.update(content)
    return digest.hexdigest()[:16]


def _normalized_root(value: Union[str, os.PathLike]) -> str:
    resolved = str(Path(value).expanduser().resolve())
    return os.path.normcase(resolved)


def build_service_identity(
    project_root_value: Union[str, os.PathLike],
    fingerprint: str,
    instance_id: Optional[str] = None,
) -> dict:
    normalized_root = _normalized_root(project_root_value)
    project_id = hashlib.sha256(normalized_root.encode("utf-8")).hexdigest()[:16]
    fingerprint = str(fingerprint)
    identity = {
        "service": SERVICE_NAME,
        "protocolVersion": PROTOCOL_VERSION,
        "projectRoot": normalized_root,
        "projectId": project_id,
        "codeFingerprint": fingerprint,
    }
    if instance_id:
        identity["instanceId"] = str(instance_id)
    return identity


def current_service_identity(instance_id: Optional[str] = None) -> dict:
    root = project_root()
    return build_service_identity(root, code_fingerprint(root), instance_id=instance_id)


def new_service_identity() -> dict:
    identity = current_service_identity(instance_id=uuid.uuid4().hex)
    launched_by_pid = os.environ.get("WPS_BRIDGE_LAUNCHED_BY_PID", "").strip()
    try:
        launched_by_pid_value = int(launched_by_pid)
    except (TypeError, ValueError):
        launched_by_pid_value = os.getppid()
    identity.update(
        {
            "launchId": os.environ.get("WPS_BRIDGE_LAUNCH_ID") or uuid.uuid4().hex,
            "launchedByPid": launched_by_pid_value,
            "serverPath": os.environ.get("WPS_BRIDGE_SERVER_PATH")
            or str(project_root() / "bridge" / "server.py"),
            "launchStartedAt": os.environ.get("WPS_BRIDGE_LAUNCH_STARTED_AT")
            or _utc_timestamp(),
        }
    )
    return identity


@dataclass(frozen=True)
class HealthInspection:
    state: str
    health: Optional[dict]
    expected: dict

    @property
    def reusable(self) -> bool:
        return self.state == "current"

    @property
    def same_project(self) -> bool:
        return bool(
            isinstance(self.health, dict)
            and self.health.get("service") == SERVICE_NAME
            and self.health.get("projectId")
            and self.health.get("projectId") == self.expected.get("projectId")
            and self.health.get("projectRoot") == self.expected.get("projectRoot")
            and self.health.get("instanceId")
        )

    @property
    def error(self) -> Optional[str]:
        if self.state == "current":
            return None
        if self.state == "absent":
            return "bridge 未运行"
        if self.state == "legacy":
            return "端口上的 bridge 缺少实例身份，可能是旧版本，拒绝复用"
        if self.state == "stale":
            return "当前 checkout 的 bridge 代码已过期，请先执行 service.py restart"
        if self.state == "foreign":
            actual_root = (self.health or {}).get("projectRoot") or "未知目录"
            return f"端口被另一份 checkout 的 bridge 占用: {actual_root}"
        if self.state == "stopping":
            return "当前 checkout 的 bridge 正在关闭"
        return "端口上的进程不是健康的 WPS bridge"


def inspect_health(health, expected: Optional[dict] = None) -> HealthInspection:
    expected = dict(expected or current_service_identity())
    if health is None:
        return HealthInspection("absent", None, expected)
    if not isinstance(health, dict) or health.get("status") != "ok":
        return HealthInspection("unhealthy", health if isinstance(health, dict) else None, expected)
    if (
        health.get("service") != SERVICE_NAME
        or not health.get("projectId")
        or not health.get("projectRoot")
        or not health.get("instanceId")
    ):
        return HealthInspection("legacy", health, expected)
    if (
        health.get("projectId") != expected.get("projectId")
        or health.get("projectRoot") != expected.get("projectRoot")
    ):
        return HealthInspection("foreign", health, expected)
    if health.get("shutdownRequested"):
        return HealthInspection("stopping", health, expected)
    if (
        health.get("protocolVersion") != expected.get("protocolVersion")
        or health.get("codeFingerprint") != expected.get("codeFingerprint")
    ):
        return HealthInspection("stale", health, expected)
    return HealthInspection("current", health, expected)


@dataclass(frozen=True)
class ServiceStartResult:
    ok: bool
    code: Optional[str] = None
    error: Optional[str] = None
    health: Optional[dict] = None
    started: bool = False
    disposition: Optional[str] = None
    listener_before: Optional[dict] = None
    listener_after: Optional[dict] = None

    def __bool__(self) -> bool:
        return self.ok


class BridgeLifecycle:
    """一次 bridge 进程的停止状态与空闲计时。"""

    def __init__(
        self,
        identity: dict,
        idle_timeout_seconds: float,
        clock: Callable[[], float] = time.monotonic,
    ):
        self.identity = dict(identity)
        self.idle_timeout_seconds = max(0.0, float(idle_timeout_seconds))
        self._clock = clock
        self._started_monotonic = clock()
        self._last_action_monotonic = self._started_monotonic
        self._started_at = _utc_timestamp()
        self._stop_event = threading.Event()
        self._stop_reason = None
        self._signal_stop_reason = None
        # Python 信号处理器可能在主线程持锁期间调用 request_stop；RLock 避免
        # SIGINT/SIGTERM 恰好落在 health/idle 临界区时自锁。
        self._lock = threading.RLock()

    @property
    def stop_reason(self) -> Optional[str]:
        with self._lock:
            return self._stop_reason

    def request_stop(self, reason: str) -> None:
        with self._lock:
            if self._stop_reason is None:
                self._stop_reason = str(reason)
            self._stop_event.set()

    def request_stop_from_signal(self, reason: str) -> None:
        """信号 handler 专用：只做原子赋值，不获取锁或重入 Event。"""
        if self._signal_stop_reason is None:
            self._signal_stop_reason = str(reason)

    def mark_action_completed(self) -> None:
        with self._lock:
            self._last_action_monotonic = self._clock()

    def idle_for_seconds(self) -> float:
        with self._lock:
            return max(0.0, self._clock() - self._last_action_monotonic)

    def should_stop(self) -> bool:
        signal_reason = self._signal_stop_reason
        if signal_reason is not None:
            self.request_stop(signal_reason)
        if self._stop_event.is_set():
            return True
        if self.idle_timeout_seconds and self.idle_for_seconds() >= self.idle_timeout_seconds:
            self.request_stop("idle_timeout")
            return True
        return False

    def health_snapshot(self) -> dict:
        server_pid = os.getpid()
        snapshot = {
            "status": "ok",
            **self.identity,
            "pid": server_pid,
            "serverPid": server_pid,
            "startedAt": self._started_at,
            "uptimeSeconds": round(max(0.0, self._clock() - self._started_monotonic), 3),
            "idleForSeconds": round(self.idle_for_seconds(), 3),
            "idleTimeoutSeconds": self.idle_timeout_seconds,
            "shutdownRequested": bool(
                self._stop_event.is_set() or self._signal_stop_reason is not None
            ),
        }
        if self.stop_reason:
            snapshot["shutdownReason"] = self.stop_reason
        return snapshot
