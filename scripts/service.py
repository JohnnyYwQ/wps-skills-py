#!/usr/bin/env python3
"""WPS bridge 生命周期命令。

用法：
    python scripts/service.py status
    python scripts/service.py start
    python scripts/service.py stop
    python scripts/service.py restart

默认只操作当前 checkout（同目录，即使代码指纹已变）的 bridge。显式
--takeover 才允许协作式关闭另一 checkout 的新版 WPS bridge；旧版或未知进程
始终拒绝强杀。
"""

from __future__ import annotations

import argparse
import errno
import json
import os
import time
import urllib.error
import urllib.request

import call
from service_lifecycle import (
    PROTOCOL_VERSION,
    SERVICE_NAME,
    current_service_identity,
    inspect_health,
)


STOP_WAIT_SECONDS = 20


def _manual_migration_hint():
    port = call.PORT
    return (
        "旧版 bridge 没有安全 shutdown 协议。先保存它持有的文档，再定位监听进程："
        f"Windows PowerShell 用 Get-NetTCPConnection -LocalPort {port} -State Listen；"
        f"macOS/Linux 用 lsof -nP -iTCP:{port} -sTCP:LISTEN。"
        "核对进程路径后再由人工终止。"
    )


def _health_probe():
    probe = call._as_health_probe(call._health(with_state=True))
    if probe.state == "responded" and probe.health is None:
        return call.HealthProbe(
            "unhealthy",
            error="端口上的 /health 返回 JSON null",
        )
    return probe


def status():
    probe = _health_probe()
    health = probe.health
    if probe.state != "responded":
        state = "stopped" if probe.state == "absent" else probe.state
        result = {
            "success": True,
            "state": state,
            "reusable": False,
            "expected": current_service_identity(),
            "service": health,
        }
        if probe.state != "absent":
            result["hint"] = (
                "bridge 端口未能完成健康检查，可能正在执行长 Action 或已卡住；"
                "不要重复启动，请稍后重试"
            )
            if probe.error:
                result["error"] = probe.error
        return result
    inspection = inspect_health(health)
    state = "stopped" if inspection.state == "absent" else inspection.state
    result = {
        "success": True,
        "state": state,
        "reusable": inspection.reusable,
        "expected": current_service_identity(),
        "service": health,
    }
    if inspection.state == "legacy":
        result["hint"] = _manual_migration_hint()
    elif inspection.error:
        result["hint"] = inspection.error
    return result


def _post_shutdown(health):
    request = urllib.request.Request(
        f"{call.BASE}/shutdown",
        data=b"{}",
        headers={
            "Content-Type": "application/json",
            "X-WPS-Bridge-Project-Id": health["projectId"],
            "X-WPS-Bridge-Instance-Id": health["instanceId"],
        },
        method="POST",
    )
    try:
        with call._urlopen(request, timeout=5) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        try:
            result = json.loads(exc.read().decode("utf-8"))
        except Exception:
            result = {"success": False, "error": str(exc)}
        result.setdefault("success", False)
        result.setdefault("httpStatus", exc.code)
        return result
    except Exception as exc:
        return {"success": False, "error": f"shutdown 请求失败: {exc}"}


def _windows_wait_result_is_running(wait_result):
    return wait_result != 0x00000000


def _windows_open_error_means_running(error_code):
    # ERROR_INVALID_PARAMETER 是查询不存在 PID 时的明确结果；拒绝访问、
    # 资源不足及未知错误都不能作为“已退出”的证据。
    return error_code != 87


def _process_is_running(pid):
    """只读检查 PID 是否仍存活；Windows 路径不会用 os.kill(0)。"""
    if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return False
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            synchronize = 0x00100000
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
            kernel32.OpenProcess.restype = wintypes.HANDLE
            kernel32.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
            kernel32.WaitForSingleObject.restype = wintypes.DWORD
            kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
            kernel32.CloseHandle.restype = wintypes.BOOL
            handle = kernel32.OpenProcess(synchronize, False, pid)
            if not handle:
                return _windows_open_error_means_running(ctypes.get_last_error())
            try:
                # 只有明确 signaled 才表示进程已退出；WAIT_TIMEOUT、
                # WAIT_FAILED 和未知值都保守视为仍存活。
                return _windows_wait_result_is_running(
                    kernel32.WaitForSingleObject(handle, 0)
                )
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return True
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError as exc:
        return exc.errno != errno.ESRCH


def _wait_until_instance_stops(instance_id, pid=None, timeout=STOP_WAIT_SECONDS):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        probe = _health_probe()
        health = probe.health
        if probe.state in ("unresponsive", "unhealthy"):
            if pid is not None and not _process_is_running(pid):
                return True
            time.sleep(0.1)
            continue
        endpoint_is_same_instance = bool(
            isinstance(health, dict) and health.get("instanceId") == instance_id
        )
        if not endpoint_is_same_instance and not _process_is_running(pid):
            return True
        time.sleep(0.1)
    return False


def stop(takeover=False):
    probe = _health_probe()
    health = probe.health
    if probe.state in ("unresponsive", "unhealthy"):
        return {
            "success": False,
            "state": probe.state,
            "code": "BRIDGE_UNAVAILABLE",
            "error": (
                "bridge 端口未能完成健康检查，可能正在执行长 Action 或已卡住；"
                "未发送停止请求，请稍后重试"
            ),
        }
    inspection = inspect_health(health)
    if inspection.state == "absent":
        return {"success": True, "state": "stopped", "alreadyStopped": True}

    if inspection.state == "stopping":
        instance_id = health["instanceId"]
        if _wait_until_instance_stops(instance_id, pid=health.get("pid")):
            return {
                "success": True,
                "state": "stopped",
                "stoppedInstanceId": instance_id,
                "alreadyStopping": True,
            }
        return {
            "success": False,
            "state": "stopping",
            "code": "BRIDGE_STOP_TIMEOUT",
            "error": f"bridge {instance_id} 在 {STOP_WAIT_SECONDS}s 内未退出",
        }

    cooperative_foreign = bool(
        takeover
        and inspection.state == "foreign"
        and isinstance(health, dict)
        and health.get("service") == SERVICE_NAME
        and health.get("protocolVersion") == PROTOCOL_VERSION
        and health.get("projectId")
        and health.get("projectRoot")
        and health.get("codeFingerprint")
        and health.get("instanceId")
        and isinstance(health.get("pid"), int)
        and not isinstance(health.get("pid"), bool)
        and health.get("pid") > 0
    )
    allowed = inspection.state == "current" or inspection.same_project or cooperative_foreign
    if not allowed:
        return {
            "success": False,
            "state": inspection.state,
            "code": "BRIDGE_STOP_REFUSED",
            "error": inspection.error,
            "hint": (
                _manual_migration_hint()
                if inspection.state == "legacy"
                else "其他 checkout 不会被强杀；确认数据安全后，可对新版 bridge 显式使用 --takeover"
            ),
            "service": health,
        }

    instance_id = health["instanceId"]
    pid = health.get("pid")
    response = _post_shutdown(health)
    if not response.get("success"):
        # health 与 shutdown 之间目标可能刚好因 idle/signal 自行退出。只有
        # transport 失败（没有服务端 code）且原 PID 已消失时才按幂等成功。
        if not response.get("code") and _wait_until_instance_stops(instance_id, pid=pid):
            return {
                "success": True,
                "state": "stopped",
                "stoppedInstanceId": instance_id,
                "shutdownRace": True,
            }
        return {
            "success": False,
            "state": inspection.state,
            "code": response.get("code") or "BRIDGE_STOP_FAILED",
            "error": response.get("error") or "bridge 拒绝停止",
            "service": health,
        }

    if not _wait_until_instance_stops(instance_id, pid=pid):
        return {
            "success": False,
            "state": "stopping",
            "code": "BRIDGE_STOP_TIMEOUT",
            "error": f"bridge {instance_id} 在 {STOP_WAIT_SECONDS}s 内未退出",
        }
    return {
        "success": True,
        "state": "stopped",
        "stoppedInstanceId": instance_id,
    }


def start():
    ready = call._ensure_server()
    return {
        "success": bool(ready),
        "state": "running" if ready else "start_failed",
        "code": ready.code,
        "error": ready.error,
        "service": ready.health,
        "started": ready.started,
        "disposition": ready.disposition,
        "listenerBefore": ready.listener_before,
        "listenerAfter": ready.listener_after,
    }


def restart(takeover=False):
    stopped = stop(takeover=takeover)
    if not stopped.get("success"):
        return stopped
    started = start()
    if started.get("success"):
        started["restarted"] = not stopped.get("alreadyStopped", False)
        started["previousInstanceId"] = stopped.get("stoppedInstanceId")
    return started


def main(argv=None):
    parser = argparse.ArgumentParser(description="管理当前 checkout 的 WPS bridge")
    parser.add_argument("command", choices=("status", "start", "stop", "restart"))
    parser.add_argument(
        "--takeover",
        action="store_true",
        help="允许协作式关闭另一 checkout 的新版 WPS bridge（绝不强杀旧版/未知进程）",
    )
    args = parser.parse_args(argv)

    if args.command == "status":
        result = status()
    elif args.command == "start":
        result = start()
    elif args.command == "stop":
        result = stop(takeover=args.takeover)
    else:
        result = restart(takeover=args.takeover)

    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("success") else 1


if __name__ == "__main__":
    raise SystemExit(main())
