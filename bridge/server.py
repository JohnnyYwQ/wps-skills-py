#!/usr/bin/env python3
"""
WPS 统一桥接服务（HTTP）

统一承载 Excel / PPT / Word 三大控制器的调度：
  - 按 action 名和显式 app 路由到对应应用控制器（Excel→Ket / PPT→Kwpp / Word→Kwps）
  - 唯一 Action 自动路由，重名 Action 缺少 app 时拒绝猜测
  - ping / wireCheck 在 server 层聚合三应用连通性

模型/用户通过 scripts/call.py 调用，无需直接发 HTTP。
HTTP handler 有界并发；Action 仍严格串行进入 controller，避免 PowerShell 单行协议交错。
"""

import json
import sys
import os
import signal
import socket
import threading
import time
from http.server import BaseHTTPRequestHandler, HTTPServer
from socketserver import ThreadingMixIn
from urllib.parse import urlparse

# 把 bridge 目录加入导入路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import wps_excel
import wps_ppt
import wps_word
from action_catalog import (
    ActionCatalog,
    ActionManifestError,
    ActionNotSupportedError,
    ActionValidationError,
    UnknownActionError,
)
from action_trace import ActionTrace
from service_lifecycle import (
    BridgeLifecycle,
    bridge_host,
    bridge_port,
    configured_idle_timeout,
    new_service_identity,
)

HOST = bridge_host()
PORT = bridge_port()

EXEC_TIMEOUT = 60
MAX_REQUEST_BYTES = 16 * 1024 * 1024
REQUEST_IO_TIMEOUT_SECONDS = 10
MAX_HTTP_HANDLER_THREADS = 16

try:
    ACTION_CATALOG = ActionCatalog.from_path()
    ACTION_MANIFEST_ERROR = None
except ActionManifestError as exc:
    ACTION_CATALOG = None
    ACTION_MANIFEST_ERROR = exc

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
    if ACTION_CATALOG is None:
        return []
    return [owner for owner in ACTION_CATALOG.owners_for(action) if owner in APP_MODULES]


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

    inferred_app = _infer_app_from_params(params)
    if inferred_app in supported_apps:
        return {
            "app": inferred_app,
            "source": "file_extension",
            "supportedApps": supported_apps,
        }, None

    default_app = ACTION_CATALOG.routing_default(action)
    if default_app in supported_apps:
        return {
            "app": default_app,
            "source": "default",
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


def _prepare_dispatch(action, params, app=None, trace=None):
    """无副作用地校验并路由 Action；不接触 controller。"""
    dispatch_started = time.perf_counter()
    if trace:
        trace.event("dispatch.started", action=action, requestedApp=app)
    if not action:
        result = {"success": False, "code": "MISSING_ACTION", "error": "缺少 action 参数"}
        if trace:
            trace.event("dispatch.rejected", status="error", code=result["code"], error=result["error"])
        return None, result

    if ACTION_MANIFEST_ERROR is not None:
        result = _route_error("INVALID_ACTION_MANIFEST", str(ACTION_MANIFEST_ERROR))
        if trace:
            trace.event("dispatch.rejected", status="error", code=result["code"], error=result["error"])
        return None, result

    if params is None:
        params = {}
    if not isinstance(params, dict):
        result = _route_error("INVALID_PARAMS", "params 必须是 JSON object")
        if trace:
            trace.event("dispatch.rejected", status="error", code=result["code"], error=result["error"])
        return None, result

    owners = ACTION_CATALOG.owners_for(action)
    if owners == ["bridge"]:
        action_params = dict(params)
        action_params.pop("app", None)
        try:
            ACTION_CATALOG.validate_params("bridge", action, action_params)
        except ActionValidationError as exc:
            result = _route_error("INVALID_PARAMS", str(exc))
            if trace:
                trace.event("dispatch.rejected", status="error", code=result["code"], error=result["error"])
            return None, result
        return {
            "action": action,
            "params": action_params,
            "requestedApp": app,
            "route": {"app": "bridge", "source": "action_registry", "supportedApps": ["bridge"]},
            "started": dispatch_started,
        }, None

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
        return None, route_error

    if trace:
        trace.event(
            "route.selected",
            action=action,
            app=route["app"],
            source=route["source"],
            supportedApps=route["supportedApps"],
        )
    action_params = dict(params)
    action_params.pop("app", None)
    try:
        ACTION_CATALOG.validate_params(route["app"], action, action_params)
    except ActionValidationError as exc:
        result = _route_error("INVALID_PARAMS", str(exc))
        if trace:
            trace.event(
                "dispatch.rejected",
                status="error",
                action=action,
                app=route["app"],
                code=result["code"],
                error=result["error"],
            )
        return None, result
    return {
        "action": action,
        "params": action_params,
        "requestedApp": app,
        "route": route,
        "started": dispatch_started,
    }, None


def dispatch(action, params, app=None, trace=None, prepared=None):
    """统一派发入口。prepared 仅供 HTTP handler 复用锁外校验结果。"""
    if prepared is None:
        prepared, rejection = _prepare_dispatch(action, params, app=app, trace=trace)
        if rejection is not None:
            return rejection

    action = prepared["action"]
    params = prepared["params"]
    route = prepared["route"]
    selected_app = route["app"]
    if selected_app == "bridge":
        result = handle_ping(trace=trace)
        try:
            ACTION_CATALOG.validate_result("bridge", action, result.get("data", {}))
        except ActionValidationError as exc:
            return _route_error("INVALID_RESULT", str(exc))
        return result
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

    execute_started = time.perf_counter()
    try:
        result = ctrl.execute(action, params, trace=trace)
        if not isinstance(result, dict):
            result = {"success": True, "data": result}
    except Exception as exc:
        result = {"success": False, "error": f"执行失败: {exc}"}

    if result.get("success"):
        try:
            ACTION_CATALOG.validate_result(
                selected_app,
                action,
                result.get("data", {}),
            )
        except ActionValidationError as exc:
            result = _route_error("INVALID_RESULT", str(exc))

    if trace:
        trace.event(
            "controller.execute.completed",
            status="success" if result.get("success") else "error",
            app=selected_app,
            action=action,
            code=result.get("code"),
            error=result.get("error"),
            elapsedMs=round((time.perf_counter() - execute_started) * 1000, 2),
            dispatchElapsedMs=round((time.perf_counter() - prepared["started"]) * 1000, 2),
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
    """Return compact public contracts from the manifest-backed registry."""
    if ACTION_MANIFEST_ERROR is not None:
        raise ACTION_MANIFEST_ERROR
    return ACTION_CATALOG.list()

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
            self._send(self.server.lifecycle.health_snapshot())
        elif parsed.path == "/actions":
            try:
                actions = get_action_list()
                self._send({"actions": actions, "count": len(actions)})
            except ActionManifestError as exc:
                self._send({"success": False, "code": exc.code, "error": str(exc)}, 500)
        elif parsed.path.startswith("/actions/"):
            if ACTION_MANIFEST_ERROR is not None:
                self._send({
                    "success": False,
                    "code": ACTION_MANIFEST_ERROR.code,
                    "error": str(ACTION_MANIFEST_ERROR),
                }, 500)
                return
            parts = parsed.path.strip("/").split("/")
            if len(parts) != 3:
                self._send({"success": False, "code": "UNKNOWN_ACTION", "error": "Action Contract 路径无效"}, 404)
                return
            _prefix, owner, action = parts
            try:
                self._send(ACTION_CATALOG.get(owner, action))
            except ActionNotSupportedError as exc:
                self._send({
                    "success": False,
                    "code": exc.code,
                    "error": str(exc),
                    "supportedApps": exc.supported_owners,
                }, 404)
            except UnknownActionError as exc:
                self._send({"success": False, "code": exc.code, "error": str(exc)}, 404)
        else:
            self._send(
                {
                    "error": "not found",
                    "routes": ["/dispatch", "/shutdown", "/health", "/actions"],
                },
                404,
            )

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/shutdown":
            self._handle_shutdown()
            return
        if parsed.path != "/dispatch":
            self._send({"error": "only /dispatch and /shutdown support POST"}, 404)
            return
        identity_error = self._validate_instance_headers()
        if identity_error:
            self._send(identity_error[0], identity_error[1])
            return
        self._handle_dispatch()

    def _validate_instance_headers(self):
        lifecycle = self.server.lifecycle
        identity = lifecycle.identity
        if lifecycle.should_stop():
            return ({
                "success": False,
                "code": "BRIDGE_SHUTTING_DOWN",
                "error": "bridge 正在关闭，请稍后重试",
            }, 503)
        requested_project = self.headers.get("X-WPS-Bridge-Project-Id")
        requested_instance = self.headers.get("X-WPS-Bridge-Instance-Id")
        if requested_project != identity.get("projectId"):
            return ({
                "success": False,
                "code": "BRIDGE_PROJECT_MISMATCH",
                "error": "请求来自另一份 checkout 或缺少项目身份",
            }, 409)
        if requested_instance != identity.get("instanceId"):
            return ({
                "success": False,
                "code": "STALE_BRIDGE_INSTANCE",
                "error": "bridge 实例已变化，请重新执行健康检查",
            }, 409)
        return None

    def _handle_shutdown(self):
        identity_error = self._validate_instance_headers()
        if identity_error:
            self._send(identity_error[0], identity_error[1])
            return
        lifecycle = self.server.lifecycle
        result = {
            "success": True,
            "status": "stopping",
            "instanceId": lifecycle.identity["instanceId"],
        }
        # 先切换为 stopping，再写 202，避免并发 dispatch 在响应写入窗口被接纳。
        # 主循环随后停止接受新连接；server_close() 等待已启动的 handler 完成。
        lifecycle.request_stop("api")
        self._send(result, 202)

    def _handle_dispatch(self):
        request_started = time.perf_counter()
        header_trace_id = self.headers.get("X-WPS-Trace-Id")
        try:
            length = int(self.headers.get("Content-Length", 0))
        except (TypeError, ValueError) as exc:
            trace = ActionTrace.resume(header_trace_id, component="server")
            self._send(trace.decorate({
                "success": False,
                "code": "INVALID_CONTENT_LENGTH",
                "error": f"Content-Length 无效: {exc}",
            }), 400)
            return
        if length < 0:
            trace = ActionTrace.resume(header_trace_id, component="server")
            self._send(trace.decorate({
                "success": False,
                "code": "INVALID_CONTENT_LENGTH",
                "error": "Content-Length 不能为负数",
            }), 400)
            return
        if length > MAX_REQUEST_BYTES:
            trace = ActionTrace.resume(header_trace_id, component="server")
            self._send(trace.decorate({
                "success": False,
                "code": "REQUEST_TOO_LARGE",
                "error": f"请求体超过 {MAX_REQUEST_BYTES} bytes 限制",
            }), 413)
            return
        try:
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

        prepared, rejection = _prepare_dispatch(action, params, app=app, trace=trace)
        if rejection is not None:
            trace.event(
                "http.response.ready",
                status="error",
                action=action,
                code=rejection.get("code"),
                error=rejection.get("error"),
                elapsedMs=round((time.perf_counter() - request_started) * 1000, 2),
            )
            self._send(trace.decorate(rejection))
            return

        lifecycle = self.server.lifecycle
        with lifecycle.action_execution(action, trace.trace_id) as admitted:
            if not admitted:
                self._send(trace.decorate({
                    "success": False,
                    "code": "BRIDGE_SHUTTING_DOWN",
                    "error": "bridge 正在关闭，请稍后重试",
                }), 503)
                return
            result = dispatch(
                action,
                params,
                app=app,
                trace=trace,
                prepared=prepared,
            )
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


class BridgeHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = False
    block_on_close = True

    # Windows 的 SO_REUSEADDR 可能允许第二监听者抢占相同端口；Unix 保留
    # reuse 以避免服务正常重启被 TIME_WAIT 阻断。
    allow_reuse_address = os.name != "nt"

    def __init__(
        self,
        server_address,
        handler_class,
        lifecycle,
        max_handler_threads=MAX_HTTP_HANDLER_THREADS,
    ):
        self.lifecycle = lifecycle
        self._handler_slots = threading.BoundedSemaphore(max_handler_threads)
        super().__init__(server_address, handler_class)

    def process_request(self, request, client_address):
        if not self._handler_slots.acquire(blocking=False):
            self.shutdown_request(request)
            return
        try:
            super().process_request(request, client_address)
        except Exception:
            self._handler_slots.release()
            raise

    def process_request_thread(self, request, client_address):
        try:
            super().process_request_thread(request, client_address)
        finally:
            self._handler_slots.release()

    def server_bind(self):
        if os.name == "nt":
            self.allow_reuse_address = False
            exclusive = getattr(socket, "SO_EXCLUSIVEADDRUSE", None)
            if exclusive is not None:
                self.socket.setsockopt(socket.SOL_SOCKET, exclusive, 1)
        super().server_bind()

    def get_request(self):
        connection, address = super().get_request()
        connection.settimeout(REQUEST_IO_TIMEOUT_SECONDS)
        return connection, address


def serve_until_stopped(httpd, lifecycle, poll_interval=0.25):
    """接受请求直至停止；handler 并发，Action 生命周期由 lifecycle 串行协调。"""
    httpd.timeout = poll_interval
    while not lifecycle.should_stop():
        httpd.handle_request()
    return lifecycle.stop_reason


def _cleanup_controllers(cache=None):
    cache = _controller_cache if cache is None else cache
    if cache is _controller_cache:
        with _controller_lock:
            controllers = list(cache.values())
            cache.clear()
    else:
        controllers = list(cache.values())
        cache.clear()
    for ctrl in controllers:
        try:
            ctrl.close()
        except Exception:
            pass


def _install_signal_handlers(lifecycle):
    previous = {}

    def request_stop(signum, _frame):
        try:
            reason = f"signal:{signal.Signals(signum).name}"
        except Exception:
            reason = f"signal:{signum}"
        lifecycle.request_stop_from_signal(reason)

    for signum in (signal.SIGINT, signal.SIGTERM):
        try:
            previous[signum] = signal.signal(signum, request_stop)
        except (ValueError, OSError):
            continue
    return previous


def _restore_signal_handlers(previous):
    for signum, handler in previous.items():
        try:
            signal.signal(signum, handler)
        except (ValueError, OSError):
            pass


def main():
    lifecycle = BridgeLifecycle(
        identity=new_service_identity(),
        idle_timeout_seconds=configured_idle_timeout(),
    )
    server = BridgeHTTPServer((HOST, PORT), Handler, lifecycle)
    previous_handlers = _install_signal_handlers(lifecycle)
    print(f"WPS 统一桥接服务已启动: http://{HOST}:{PORT}", flush=True)
    counts = {
        owner: len(ACTION_CATALOG.list(owner=owner))
        for owner in ("excel", "ppt", "word")
    }
    print(
        f"  支持 contract: Excel {counts['excel']} + PPT {counts['ppt']} + Word {counts['word']}",
        flush=True,
    )
    print(
        f"  PID: {os.getpid()} | instance: {lifecycle.identity['instanceId']}"
        f" | idle timeout: {lifecycle.idle_timeout_seconds:g}s",
        flush=True,
    )
    print("  按 Ctrl+C 或运行 python scripts/service.py stop 停止", flush=True)
    try:
        reason = serve_until_stopped(server, lifecycle)
        print(f"bridge 正在停止: {reason}", flush=True)
    except KeyboardInterrupt:
        lifecycle.request_stop("keyboard_interrupt")
    finally:
        try:
            # 先释放监听端口，阻止清理阶段继续积压请求；service.py 会继续按
            # 原 PID 等待，直到 controller 清理完成、进程真正退出。
            server.server_close()
        finally:
            try:
                _cleanup_controllers()
            finally:
                # 清理期间继续保留安全的信号 handler，避免第二次 SIGTERM
                # 直接打断 PowerShell/COM 的有界释放流程。
                _restore_signal_handlers(previous_handlers)
        print("服务已停止", flush=True)

if __name__ == "__main__":
    main()
