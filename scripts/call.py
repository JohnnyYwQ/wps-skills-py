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

import sys
import os
import json
import time
import subprocess
import urllib.request
import urllib.error

HOST = "127.0.0.1"
PORT = 58891
BASE = f"http://{HOST}:{PORT}"

BRIDGE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "bridge")
if BRIDGE_DIR not in sys.path:
    sys.path.insert(0, BRIDGE_DIR)

from action_trace import ActionTrace, server_log_path


def _project_root():
    """scripts/ 的上级目录 = 技能根目录"""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _server_path():
    return os.path.join(_project_root(), "bridge", "server.py")


def _health():
    try:
        with urllib.request.urlopen(f"{BASE}/health", timeout=2) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def _is_healthy(h):
    """桥接服务 /health 返回 {"status":"ok"}；兼容旧版 {"success":true}。"""
    return bool(h and (h.get("status") == "ok" or h.get("success")))


def _ensure_server(trace=None):
    """确保桥接服务在运行；未运行则后台拉起。返回是否成功。"""
    started = time.perf_counter()
    if _is_healthy(_health()):
        if trace:
            trace.event(
                "bridge.health.checked",
                status="healthy",
                elapsedMs=round((time.perf_counter() - started) * 1000, 2),
            )
        return True

    py = sys.executable
    server = _server_path()
    if not os.path.exists(server):
        if trace:
            trace.event("bridge.start.failed", status="error", error=f"找不到桥接服务: {server}")
        return False

    log_path, log_warning = server_log_path()
    if trace:
        trace.event(
            "bridge.start.requested",
            serverPath=server,
            serverLog=str(log_path) if log_path else None,
        )
        if log_warning:
            trace.event("bridge.log.warning", status="warning", warning=log_warning)

    log_stream = None
    try:
        if log_path is not None:
            log_stream = log_path.open("a", encoding="utf-8")
        stdout_target = log_stream if log_stream is not None else subprocess.DEVNULL
        stderr_target = subprocess.STDOUT if log_stream is not None else subprocess.DEVNULL
        if os.name == "nt":
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            subprocess.Popen(
                [py, server],
                stdout=stdout_target,
                stderr=stderr_target,
                creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
                startupinfo=si,
            )
        else:
            subprocess.Popen(
                [py, server],
                stdout=stdout_target,
                stderr=stderr_target,
                start_new_session=True,
            )
    except Exception as e:
        if trace:
            trace.event("bridge.start.failed", status="error", error=f"启动桥接服务失败: {e}")
        return False
    finally:
        if log_stream is not None:
            log_stream.close()

    # 轮询等待服务就绪（最多 ~10s）
    for _ in range(40):
        time.sleep(0.25)
        if _is_healthy(_health()):
            if trace:
                trace.event(
                    "bridge.start.completed",
                    status="healthy",
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )
            return True
    if trace:
        trace.event(
            "bridge.start.failed",
            status="timeout",
            elapsedMs=round((time.perf_counter() - started) * 1000, 2),
            serverLog=str(log_path) if log_path else None,
        )
    return False


def _post(action, params, app=None, trace=None):
    trace = trace or ActionTrace.start(component="call")
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
        with urllib.request.urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            http_status = getattr(resp, "status", 200)
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

    if not _ensure_server(trace=trace):
        result = {
            "success": False,
            "error": "桥接服务启动失败，请检查 traceLog 与 server 日志",
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
        result = _post(action, params, app=app, trace=trace)
    except urllib.error.URLError as e:
        result = {"success": False, "error": f"调用失败（服务无响应）: {e}"}
        trace.event(
            "action.completed",
            status="error",
            action=action,
            elapsedMs=round((time.perf_counter() - action_started) * 1000, 2),
        )
        print(json.dumps(trace.decorate(result), ensure_ascii=False))
        sys.exit(1)
    except Exception as e:
        result = {"success": False, "error": f"调用失败: {e}"}
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
