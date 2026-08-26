#!/usr/bin/env python3
"""
WPS 统一桥接服务（HTTP）

统一承载 Excel / PPT / Word 三大控制器的调度：
  - 按 action 名和显式 app 路由到对应应用控制器（Excel→Ket / PPT→Kwpp / Word→Kwps）
  - 唯一 Action 自动路由，重名 Action 缺少 app 时拒绝猜测
  - ping / wireCheck 在 server 层聚合三应用连通性

模型/用户通过 scripts/call.py 调用，无需直接发 HTTP。
单线程 HTTPServer：底层 PowerShell 子进程是单 line 协议，多线程会交错写 stdin 导致协议崩。
"""

import json
import re
import sys
import os
import threading
import time
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse

# 把 bridge 目录加入导入路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import wps_excel
import wps_ppt
import wps_word
from action_trace import ActionTrace

HOST = "127.0.0.1"
PORT = 58891

EXEC_TIMEOUT = 60

# 从各控制器内嵌的 PS 脚本中解析出支持的 action 列表
def _actions_from_module(mod):
    return set(re.findall(r"function Exec-(\w+)\b", mod.PS_BRIDGE_SCRIPT))

EXCEL_ACTIONS = _actions_from_module(wps_excel)
PPT_ACTIONS = _actions_from_module(wps_ppt)
WORD_ACTIONS = _actions_from_module(wps_word)

# 通用 action：优先按显式 app 委派，其次按目标扩展名推断，最后兼容性回退 excel
COMMON_ACTIONS = {
    "save", "saveAs", "convertToPDF", "convertFormat",
    "getSelectedText", "setSelectedText", "getAppInfo", "reconnect",
}
SERVER_ACTIONS = {"ping", "wireCheck"}

def _infer_app_from_params(params):
    """从 filePath / outputPath / targetFormat 扩展名推断应用，避免 saveAs 无 app 时
    静默把 Excel 另存为 .pptx（假成功）。无法推断时返回 None。"""
    params = params or {}
    for key in ("filePath", "outputPath", "path"):
        v = params.get(key)
        if not v:
            continue
        v = str(v).lower()
        if v.endswith(".pptx") or v.endswith(".ppt"):
            return "ppt"
        if v.endswith(".xlsx") or v.endswith(".xls"):
            return "excel"
        if v.endswith(".docx") or v.endswith(".doc"):
            return "word"
    return None

# 应用 -> 控制器模块映射（懒加载，避免一次拉起三个 WPS）
APP_MODULES = {
    "excel": wps_excel,
    "ppt": wps_ppt,
    "word": wps_word,
}

APP_ACTIONS = {
    "excel": EXCEL_ACTIONS,
    "ppt": PPT_ACTIONS,
    "word": WORD_ACTIONS,
}

APP_PROGIDS = {
    "excel": wps_excel.EXCEL_PROGID,
    "ppt": wps_ppt.PPT_PROGID,
    "word": wps_word.WORD_PROGID,
}

_controller_cache = {}
_controller_lock = threading.Lock()

def get_app_controller(app, trace=None):
    """懒加载并缓存指定应用的控制器单例"""
    app = (app or "excel").lower()
    if app not in APP_MODULES:
        app = "excel"
    with _controller_lock:
        ctrl = _controller_cache.get(app)
        if ctrl is None or not getattr(ctrl, "_ready", False):
            ctrl = APP_MODULES[app].get_controller(trace=trace)
            _controller_cache[app] = ctrl
    return ctrl

def _route_error(code, message, supported_apps=None):
    result = {"success": False, "code": code, "error": message}
    if supported_apps is not None:
        result["supportedApps"] = list(supported_apps)
    return result


def _normalize_app(app):
    if app is None:
        return None
    normalized = str(app).strip().lower()
    return normalized or None


def _supported_apps(action):
    return [app for app, actions in APP_ACTIONS.items() if action in actions]


def route_action(action, params=None, requested_app=None):
    """解析 Action 的目标应用，返回 (route, error)。

    唯一 Action 自动路由；重名 Action 必须显式指定 app。顶层 app 优先用于
    HTTP/CLI 契约，同时兼容历史 params.app。路由字段不会传入应用控制器。
    """
    params = params or {}
    param_app = _normalize_app(params.get("app"))
    top_level_app = _normalize_app(requested_app)

    if top_level_app and param_app and top_level_app != param_app:
        return None, _route_error(
            "CONFLICTING_APP",
            f"顶层 app '{top_level_app}' 与 params.app '{param_app}' 不一致",
        )

    explicit_app = top_level_app or param_app
    if explicit_app and explicit_app not in APP_MODULES:
        return None, _route_error(
            "INVALID_APP",
            f"未知应用: {explicit_app}",
            APP_MODULES.keys(),
        )

    if action in COMMON_ACTIONS:
        if explicit_app:
            return {
                "app": explicit_app,
                "source": "explicit",
                "supportedApps": list(APP_MODULES),
            }, None

        inferred_app = _infer_app_from_params(params)
        if inferred_app:
            return {
                "app": inferred_app,
                "source": "file_extension",
                "supportedApps": list(APP_MODULES),
            }, None

        return {
            "app": "excel",
            "source": "default",
            "supportedApps": list(APP_MODULES),
        }, None

    supported_apps = _supported_apps(action)
    if not supported_apps:
        return None, _route_error(
            "UNKNOWN_ACTION",
            f"未知 action: {action}（可用 action 见 /actions）",
        )

    if explicit_app:
        if explicit_app not in supported_apps:
            return None, _route_error(
                "ACTION_NOT_SUPPORTED_FOR_APP",
                f"应用 '{explicit_app}' 不支持 action '{action}'",
                supported_apps,
            )
        return {
            "app": explicit_app,
            "source": "explicit",
            "supportedApps": supported_apps,
        }, None

    if len(supported_apps) > 1:
        return None, _route_error(
            "AMBIGUOUS_ACTION",
            f"action '{action}' 同时属于多个应用，必须显式指定 app",
            supported_apps,
        )

    return {
        "app": supported_apps[0],
        "source": "action_registry",
        "supportedApps": supported_apps,
    }, None


def dispatch(action, params, app=None, trace=None):
    """统一派发入口"""
    dispatch_started = time.perf_counter()
    if trace:
        trace.event("dispatch.started", action=action, requestedApp=app)
    if not action:
        result = {"success": False, "code": "MISSING_ACTION", "error": "缺少 action 参数"}
        if trace:
            trace.event("dispatch.rejected", status="error", code=result["code"], error=result["error"])
        return result

    if params is None:
        params = {}
    if not isinstance(params, dict):
        result = _route_error("INVALID_PARAMS", "params 必须是 JSON object")
        if trace:
            trace.event("dispatch.rejected", status="error", code=result["code"], error=result["error"])
        return result

    if action == "ping":
        return handle_ping(trace=trace)
    if action == "wireCheck":
        return handle_ping(trace=trace)  # wireCheck 与 ping 同源：检测桥接线路

    route, route_error = route_action(action, params, requested_app=app)
    if route_error:
        if trace:
            trace.event(
                "route.rejected",
                status="error",
                action=action,
                code=route_error.get("code"),
                error=route_error.get("error"),
                supportedApps=route_error.get("supportedApps"),
            )
        return route_error

    selected_app = route["app"]
    if trace:
        trace.event(
            "route.selected",
            action=action,
            app=selected_app,
            source=route["source"],
            supportedApps=route["supportedApps"],
        )

    controller_started = time.perf_counter()
    try:
        ctrl = get_app_controller(selected_app, trace=trace)
    except Exception as exc:
        result = {
            "success": False,
            "code": "CONTROLLER_INIT_FAILED",
            "error": f"{selected_app} 控制器初始化失败: {exc}",
        }
        if trace:
            trace.event(
                "controller.init.failed",
                status="error",
                app=selected_app,
                controller=f"Wps{selected_app.title()}Controller",
                error=f"{type(exc).__name__}: {exc}",
                elapsedMs=round((time.perf_counter() - controller_started) * 1000, 2),
            )
        return result

    platform_name = getattr(ctrl, "platform", "unknown")
    if platform_name == "Windows":
        backend_kind = "powershell_com"
    elif platform_name == "Linux":
        backend_kind = "openxml_file"
    else:
        backend_kind = "unsupported"
    if trace:
        trace.event(
            "controller.ready",
            app=selected_app,
            controller=type(ctrl).__name__,
            platform=platform_name,
            backend=backend_kind,
            progId=APP_PROGIDS[selected_app],
            elapsedMs=round((time.perf_counter() - controller_started) * 1000, 2),
        )

    action_params = dict(params)
    action_params.pop("app", None)

    execute_started = time.perf_counter()
    try:
        result = ctrl.execute(action, action_params, trace=trace)
        if not isinstance(result, dict):
            result = {"success": True, "data": result}
    except Exception as exc:
        result = {"success": False, "error": f"执行失败: {exc}"}

    if trace:
        trace.event(
            "controller.execute.completed",
            status="success" if result.get("success") else "error",
            app=selected_app,
            action=action,
            code=result.get("code"),
            error=result.get("error"),
            elapsedMs=round((time.perf_counter() - execute_started) * 1000, 2),
            dispatchElapsedMs=round((time.perf_counter() - dispatch_started) * 1000, 2),
            **trace.debug_fields(response=result),
        )
    return result

def handle_ping(trace=None):
    """聚合三应用连通性"""
    result = {"success": True, "data": {}}
    for app in ("excel", "ppt", "word"):
        started = time.perf_counter()
        try:
            ctrl = get_app_controller(app, trace=trace)
            ok = ctrl.ping(trace=trace)
            result["data"][app] = ok
            if not ok:
                result["success"] = True  # 整体仍返回成功，detail 里标明哪些未连
            if trace:
                trace.event(
                    "ping.app.completed",
                    app=app,
                    status="success" if ok else "unavailable",
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )
        except Exception as e:
            result["data"][app] = False
            result["data"][f"{app}_error"] = str(e)
            if trace:
                trace.event(
                    "ping.app.completed",
                    app=app,
                    status="error",
                    error=f"{type(e).__name__}: {e}",
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )
    return result

def get_action_list():
    """返回全部可用 action 及其所属应用，供 /actions 与 SKILL.md 对齐"""
    out = []
    for a in sorted(EXCEL_ACTIONS):
        if a not in COMMON_ACTIONS and a not in SERVER_ACTIONS:
            out.append({"action": a, "app": "excel"})
    for a in sorted(PPT_ACTIONS):
        if a not in COMMON_ACTIONS and a not in SERVER_ACTIONS:
            out.append({"action": a, "app": "ppt"})
    for a in sorted(WORD_ACTIONS):
        if a not in COMMON_ACTIONS and a not in SERVER_ACTIONS:
            out.append({"action": a, "app": "word"})
    for a in sorted(COMMON_ACTIONS | SERVER_ACTIONS):
        out.append({"action": a, "app": "common"})
    return out

class Handler(BaseHTTPRequestHandler):
    def _send(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self._send({"status": "ok"})
        elif parsed.path == "/actions":
            actions = get_action_list()
            self._send({"actions": actions, "count": len(actions)})
        else:
            self._send({"error": "not found", "routes": ["/dispatch", "/health", "/actions"]}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/dispatch":
            self._send({"error": "only /dispatch supports POST"}, 404)
            return
        request_started = time.perf_counter()
        header_trace_id = self.headers.get("X-WPS-Trace-Id")
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            payload = json.loads(raw.decode("utf-8") or "{}")
        except Exception as e:
            trace = ActionTrace.resume(header_trace_id, component="server")
            trace.event(
                "http.request.rejected",
                status="error",
                code="INVALID_JSON",
                error=f"{type(e).__name__}: {e}",
            )
            result = trace.decorate({
                "success": False,
                "code": "INVALID_JSON",
                "error": f"请求体解析失败: {e}",
            })
            self._send(result, 400)
            return

        if not isinstance(payload, dict):
            trace = ActionTrace.resume(header_trace_id, component="server")
            trace.event(
                "http.request.rejected",
                status="error",
                code="INVALID_PAYLOAD",
                error="请求体必须是 JSON object",
            )
            self._send(trace.decorate({
                "success": False,
                "code": "INVALID_PAYLOAD",
                "error": "请求体必须是 JSON object",
            }), 400)
            return

        body_trace_id = payload.get("traceId")
        trace = ActionTrace.resume(header_trace_id or body_trace_id, component="server")
        action = payload.get("action")
        params = payload.get("params", {})
        app = payload.get("app")
        trace.event(
            "http.request.received",
            method="POST",
            path="/dispatch",
            bodyBytes=len(raw),
            action=action,
            requestedApp=app,
            **trace.debug_fields(params=params),
        )
        if header_trace_id and body_trace_id and header_trace_id != body_trace_id:
            trace.event(
                "http.trace_id.conflict",
                status="warning",
                headerTraceId=header_trace_id,
                bodyTraceId=body_trace_id,
            )

        result = dispatch(action, params, app=app, trace=trace)
        trace.event(
            "http.response.ready",
            status="success" if result.get("success") else "error",
            action=action,
            code=result.get("code"),
            error=result.get("error"),
            elapsedMs=round((time.perf_counter() - request_started) * 1000, 2),
        )
        self._send(trace.decorate(result))

    def log_message(self, *args):
        pass  # 静默

def main():
    server = HTTPServer((HOST, PORT), Handler)
    print(f"WPS 统一桥接服务已启动: http://{HOST}:{PORT}", flush=True)
    print(f"  支持 action: Excel {len(EXCEL_ACTIONS)} + PPT {len(PPT_ACTIONS)} + Word {len(WORD_ACTIONS)}", flush=True)
    print("  按 Ctrl+C 停止", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
    finally:
        # 清理所有控制器（关闭 PS 进程）
        with _controller_lock:
            for ctrl in _controller_cache.values():
                try:
                    ctrl.close()
                except Exception:
                    pass
        server.server_close()

if __name__ == "__main__":
    main()
