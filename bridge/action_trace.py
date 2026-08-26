#!/usr/bin/env python3
"""Action 级结构化追踪。

公开接口刻意保持很小：创建/接续 ActionTrace、写事件、生成 debug 摘要、
装饰响应。日志目录选择、降级、JSONL、内容脱敏和 1 天保留都封装在模块内。
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import re
import tempfile
import threading
import time
from datetime import datetime, timezone
from typing import Optional, Tuple
import uuid


RETENTION_SECONDS = 24 * 60 * 60
CLEANUP_INTERVAL_SECONDS = 60 * 60

# ==================== 手动 trace 开关 ====================
# "info"：只记录低敏元数据、耗时和错误（默认）
# "debug"：额外记录脱敏后的调用参数和响应摘要
#
# LLM 调用 scripts/call.py 时无需改变命令；手动修改这一行即可切换。
# 若进程环境显式设置了 WPS_TRACE，则环境变量优先，便于临时覆盖。
TRACE_LEVEL = "debug"

_TRACE_ID_RE = re.compile(r"^act-(\d{8})-([a-f0-9]{32})$")
_WRITE_LOCK = threading.Lock()

_SECRET_KEY_PARTS = (
    "authorization",
    "password",
    "passwd",
    "secret",
    "token",
    "api_key",
    "apikey",
    "credential",
    "cookie",
)
_CONTENT_KEYS = {
    "body",
    "content",
    "data",
    "find",
    "find_text",
    "formula",
    "html",
    "notes",
    "replace",
    "replace_text",
    "subtitle",
    "text",
    "title",
    "value",
    "values",
    "xml",
}


def _configured_trace_level() -> str:
    """解析 trace 级别；无效值安全降级为 info，不能阻断 Action。"""
    environment_level = os.environ.get("WPS_TRACE", "").strip().lower()
    configured_level = environment_level or str(TRACE_LEVEL).strip().lower()
    return "debug" if configured_level == "debug" else "info"


def _utc_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def _new_trace_id() -> str:
    local_day = datetime.now().strftime("%Y%m%d")
    return f"act-{local_day}-{uuid.uuid4().hex}"


def _valid_trace_id(trace_id) -> bool:
    return bool(trace_id and _TRACE_ID_RE.fullmatch(str(trace_id)))


def _trace_day(trace_id: str) -> str:
    match = _TRACE_ID_RE.fullmatch(trace_id)
    day = match.group(1) if match else datetime.now().strftime("%Y%m%d")
    return f"{day[:4]}-{day[4:6]}-{day[6:8]}"


def _skill_log_root() -> Path:
    return Path(__file__).resolve().parent.parent / "logs"


def _fallback_log_root() -> Path:
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA")
        if base:
            return Path(base) / "wps-skills" / "logs"
    state_home = os.environ.get("XDG_STATE_HOME")
    if state_home:
        return Path(state_home) / "wps-skills"
    try:
        return Path.home() / ".local" / "state" / "wps-skills"
    except Exception:
        return Path(tempfile.gettempdir()) / "wps-skills-logs"


def _check_writable(root: Path) -> None:
    traces = root / "traces"
    traces.mkdir(parents=True, exist_ok=True)
    probe = traces / f".write-test-{os.getpid()}-{uuid.uuid4().hex}"
    try:
        probe.write_text("ok", encoding="utf-8")
    finally:
        try:
            probe.unlink()
        except FileNotFoundError:
            pass


def _select_log_root() -> Tuple[Optional[Path], Optional[str]]:
    override = os.environ.get("WPS_TRACE_DIR")
    candidates = []
    if override:
        candidates.append((Path(override).expanduser(), "WPS_TRACE_DIR"))
    candidates.append((_skill_log_root(), "skill"))
    fallback = _fallback_log_root()
    if all(candidate != fallback for candidate, _ in candidates):
        candidates.append((fallback, "fallback"))

    failures = []
    for candidate, source in candidates:
        try:
            candidate = candidate.resolve()
            _check_writable(candidate)
            warning = None
            if failures:
                warning = "；".join(failures) + f"；已降级到 {candidate}"
            return candidate, warning
        except Exception as exc:
            failures.append(f"{source} 日志目录不可写: {type(exc).__name__}: {exc}")
    return None, "；".join(failures) or "没有可写的 trace 日志目录"


def _remove_old_logs_unchecked(root: Path) -> None:
    traces = root / "traces"
    traces.mkdir(parents=True, exist_ok=True)
    marker = traces / ".last-cleanup"
    now = time.time()
    try:
        if marker.exists() and marker.stat().st_mtime >= now - CLEANUP_INTERVAL_SECONDS:
            return
    except OSError:
        pass

    cutoff = now - RETENTION_SECONDS
    try:
        for path in traces.rglob("*.jsonl"):
            try:
                if path.is_file() and not path.is_symlink() and path.stat().st_mtime < cutoff:
                    path.unlink()
            except (FileNotFoundError, OSError):
                continue

        directories = sorted(
            (path for path in traces.rglob("*") if path.is_dir()),
            key=lambda path: len(path.parts),
            reverse=True,
        )
        for directory in directories:
            try:
                directory.rmdir()
            except OSError:
                pass

        for path in root.glob("server-*.log"):
            try:
                if path.is_file() and not path.is_symlink() and path.stat().st_mtime < cutoff:
                    path.unlink()
            except (FileNotFoundError, OSError):
                continue
    finally:
        try:
            marker.touch()
        except OSError:
            pass


def _remove_old_logs(root: Path) -> Optional[str]:
    """惰性清理过期日志；清理失败不得影响 Action 执行。"""
    try:
        _remove_old_logs_unchecked(root)
        return None
    except Exception as exc:
        return f"过期日志清理失败: {type(exc).__name__}: {exc}"


def server_log_path() -> Tuple[Optional[Path], Optional[str]]:
    """返回后台桥接服务的普通 stdout/stderr 日志路径。"""
    root, warning = _select_log_root()
    if root is None:
        return None, warning
    cleanup_warning = _remove_old_logs(root)
    if cleanup_warning:
        warning = "；".join(part for part in (warning, cleanup_warning) if part)
    day = datetime.now().strftime("%Y-%m-%d")
    return root / f"server-{day}.log", warning


def _truncate(value: str, limit: int) -> str:
    if len(value) <= limit:
        return value
    return value[:limit] + f"…<{len(value) - limit} chars omitted>"


def _content_summary(value: str) -> dict:
    encoded = value.encode("utf-8", errors="replace")
    return {
        "type": "str",
        "length": len(value),
        "sha256": hashlib.sha256(encoded).hexdigest()[:16],
    }


def _debug_summary(value, key: Optional[str] = None, depth: int = 0):
    key_lower = (key or "").lower()
    if any(part in key_lower for part in _SECRET_KEY_PARTS):
        return "<redacted>"
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        if key_lower in _CONTENT_KEYS or key_lower.endswith("_text"):
            return _content_summary(value)
        if "path" in key_lower or key_lower.endswith("file"):
            return _truncate(value, 1024)
        return _truncate(value, 256)
    if isinstance(value, bytes):
        return {
            "type": "bytes",
            "length": len(value),
            "sha256": hashlib.sha256(value).hexdigest()[:16],
        }
    if depth >= 4:
        return {"type": type(value).__name__, "summary": "<max depth>"}
    if isinstance(value, dict):
        result = {}
        for index, (child_key, child_value) in enumerate(value.items()):
            if index >= 50:
                result["<omitted>"] = len(value) - 50
                break
            child_key = str(child_key)
            result[child_key] = _debug_summary(child_value, child_key, depth + 1)
        return result
    if isinstance(value, (list, tuple, set)):
        items = list(value)
        return {
            "type": type(value).__name__,
            "length": len(items),
            "sample": [_debug_summary(item, key, depth + 1) for item in items[:5]],
        }
    return {"type": type(value).__name__}


def _json_safe(value, depth: int = 0):
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        return _truncate(value, 4096)
    if depth >= 5:
        return f"<{type(value).__name__}: max depth>"
    if isinstance(value, dict):
        return {str(key): _json_safe(child, depth + 1) for key, child in list(value.items())[:100]}
    if isinstance(value, (list, tuple, set)):
        return [_json_safe(child, depth + 1) for child in list(value)[:100]]
    if isinstance(value, bytes):
        return {"type": "bytes", "length": len(value)}
    return f"<{type(value).__name__}>"


class ActionTrace:
    """一个 Action 的追加式 JSONL trace。所有写入失败都会降级为 warning。"""

    def __init__(self, trace_id: str, component: str):
        self.trace_id = trace_id
        self.component = component
        self.trace_level = _configured_trace_level()
        self.debug_enabled = self.trace_level == "debug"
        self._warnings = []
        root, warning = _select_log_root()
        if warning:
            self._warnings.append(warning)
        self.log_path = None
        if root is not None:
            cleanup_warning = _remove_old_logs(root)
            if cleanup_warning:
                self._warnings.append(cleanup_warning)
            self.log_path = root / "traces" / _trace_day(trace_id) / f"{trace_id}.jsonl"
            try:
                self.log_path.parent.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                self._warnings.append(f"trace 目录创建失败: {type(exc).__name__}: {exc}")
                self.log_path = None

    @classmethod
    def start(cls, component: str) -> "ActionTrace":
        return cls(_new_trace_id(), component)

    @classmethod
    def resume(cls, trace_id: Optional[str], component: str) -> "ActionTrace":
        if _valid_trace_id(trace_id):
            return cls(str(trace_id), component)
        trace = cls(_new_trace_id(), component)
        if trace_id:
            trace._warnings.append("收到无效 traceId，已重新生成")
        return trace

    @property
    def warning(self) -> Optional[str]:
        if not self._warnings:
            return None
        return "；".join(dict.fromkeys(self._warnings))

    def debug_fields(self, **fields) -> dict:
        if not self.debug_enabled:
            return {}
        return {key: _debug_summary(value, key) for key, value in fields.items()}

    def event(self, event: str, **fields) -> bool:
        if self.log_path is None:
            return False
        row = {
            "ts": _utc_timestamp(),
            "traceId": self.trace_id,
            "event": event,
            "component": self.component,
            "pid": os.getpid(),
            "thread": threading.current_thread().name,
        }
        for key, value in fields.items():
            if key not in row:
                row[key] = _json_safe(value)
        encoded = json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n"
        try:
            with _WRITE_LOCK:
                with self.log_path.open("a", encoding="utf-8", newline="") as stream:
                    stream.write(encoded)
                    stream.flush()
            return True
        except Exception as exc:
            self._warnings.append(f"trace 写入失败: {type(exc).__name__}: {exc}")
            return False

    def decorate(self, result: dict) -> dict:
        decorated = dict(result)
        decorated["traceId"] = self.trace_id
        decorated["traceLog"] = str(self.log_path) if self.log_path is not None else None
        if self.warning:
            decorated["traceWarning"] = self.warning
        return decorated
