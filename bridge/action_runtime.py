"""In-process Action execution behind the ``execute`` / ``close`` seam."""

from dataclasses import dataclass, field
from pathlib import Path
import sys
import threading
import time
from typing import Any, Callable, Mapping, Optional

from action_catalog import ActionCatalog, ActionManifestError, ActionValidationError
from action_trace import ActionTrace
import wps_excel
import wps_ppt
import wps_word


_APP_MODULES = {
    "excel": wps_excel,
    "ppt": wps_ppt,
    "word": wps_word,
}

_APP_PROGIDS = {
    "excel": wps_excel.EXCEL_PROGID,
    "ppt": wps_ppt.PPT_PROGID,
    "word": wps_word.WORD_PROGID,
}


@dataclass(frozen=True)
class ActionRequest:
    action: str
    params: Any = field(default_factory=dict)
    app: Optional[str] = None
    trace_id: Optional[str] = None


@dataclass(frozen=True)
class ActionResponse:
    """Transport-neutral result of one Action execution."""

    _body: Mapping[str, Any]
    trace_id: str
    trace_log: Optional[Path]

    @property
    def success(self) -> bool:
        return bool(self._body.get("success"))

    def to_dict(self) -> dict:
        return dict(self._body)


def _default_controller_factory(app, trace=None):
    return _APP_MODULES[app].get_controller(trace=trace)


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


class ActionRuntime:
    """Resolve, execute, trace, cache, and close local WPS Actions."""

    def __init__(
        self,
        *,
        catalog=None,
        controller_factory: Optional[Callable[..., Any]] = None,
    ):
        self._manifest_error = None
        if catalog is not None:
            self._catalog = catalog
        else:
            try:
                self._catalog = ActionCatalog.from_path()
            except ActionManifestError as exc:
                self._catalog = None
                self._manifest_error = exc
        self._controller_factory = controller_factory or _default_controller_factory
        self._controllers = {}
        self._controller_lock = threading.Lock()
        self._execution_lock = threading.Lock()
        self._state_condition = threading.Condition()
        self._state = "open"

    @staticmethod
    def _complete(trace, started, action, result):
        trace.event(
            "action.completed",
            status="success" if result.get("success") else "error",
            action=action,
            code=result.get("code"),
            error=result.get("error"),
            elapsedMs=round((time.perf_counter() - started) * 1000, 2),
        )
        decorated = trace.decorate(result)
        return ActionResponse(decorated, trace.trace_id, trace.log_path)

    def _resolve_route(self, action, params, requested_app):
        owners = self._catalog.owners_for(action)
        param_app = _normalize_app(params.get("app"))
        top_level_app = _normalize_app(requested_app)
        if top_level_app and param_app and top_level_app != param_app:
            return None, None, _route_error(
                "CONFLICTING_APP",
                f"顶层 app '{top_level_app}' 与 params.app '{param_app}' 不一致",
            )

        explicit_app = top_level_app or param_app
        if owners == ["bridge"]:
            if explicit_app and explicit_app != "bridge":
                return None, None, _route_error(
                    "ACTION_NOT_SUPPORTED_FOR_APP",
                    f"应用 '{explicit_app}' 不支持 action '{action}'",
                    ["bridge"],
                )
            action_params = dict(params)
            action_params.pop("app", None)
            return {
                "app": "bridge",
                "source": "action_registry",
                "supportedApps": ["bridge"],
            }, action_params, None

        if explicit_app and explicit_app not in _APP_MODULES:
            return None, None, _route_error(
                "INVALID_APP",
                f"未知应用: {explicit_app}",
                _APP_MODULES.keys(),
            )

        supported_apps = [owner for owner in owners if owner in _APP_MODULES]
        if not supported_apps:
            return None, None, _route_error(
                "UNKNOWN_ACTION",
                f"未知 action: {action}",
            )
        if explicit_app:
            if explicit_app not in supported_apps:
                return None, None, _route_error(
                    "ACTION_NOT_SUPPORTED_FOR_APP",
                    f"应用 '{explicit_app}' 不支持 action '{action}'",
                    supported_apps,
                )
            route = {
                "app": explicit_app,
                "source": "explicit",
                "supportedApps": supported_apps,
            }
        elif len(supported_apps) > 1:
            return None, None, _route_error(
                "AMBIGUOUS_ACTION",
                f"action '{action}' 同时属于多个应用，必须显式指定 app",
                supported_apps,
            )
        else:
            route = {
                "app": supported_apps[0],
                "source": "action_registry",
                "supportedApps": supported_apps,
            }

        action_params = dict(params)
        action_params.pop("app", None)
        return route, action_params, None

    def execute(self, request: ActionRequest) -> ActionResponse:
        trace = ActionTrace.resume(request.trace_id, component="runtime")
        started = time.perf_counter()
        trace.event(
            "action.started",
            action=request.action,
            requestedApp=request.app,
            **trace.debug_fields(params=request.params),
        )
        with self._state_condition:
            runtime_open = self._state == "open"
        if not runtime_open:
            result = {
                "success": False,
                "code": "RUNTIME_CLOSED",
                "error": "Action Runtime 已关闭",
            }
            trace.event(
                "action.completed",
                status="error",
                action=request.action,
                code=result["code"],
                error=result["error"],
                elapsedMs=round((time.perf_counter() - started) * 1000, 2),
            )
            decorated = trace.decorate(result)
            return ActionResponse(decorated, trace.trace_id, trace.log_path)

        if self._manifest_error is not None:
            result = _route_error(
                "INVALID_ACTION_MANIFEST",
                str(self._manifest_error),
            )
            trace.event(
                "dispatch.rejected",
                status="error",
                code=result["code"],
                error=result["error"],
            )
            return self._complete(trace, started, request.action, result)

        if not request.action:
            result = {
                "success": False,
                "code": "MISSING_ACTION",
                "error": "缺少 action 参数",
            }
            trace.event(
                "dispatch.rejected",
                status="error",
                code=result["code"],
                error=result["error"],
            )
            return self._complete(trace, started, request.action, result)

        params = request.params
        if not isinstance(params, dict):
            result = _route_error("INVALID_PARAMS", "params 必须是 JSON object")
            trace.event(
                "dispatch.rejected",
                status="error",
                code=result["code"],
                error=result["error"],
            )
            return self._complete(trace, started, request.action, result)

        route, action_params, route_error = self._resolve_route(
            request.action,
            params,
            request.app,
        )
        if route_error is not None:
            trace.event(
                "route.rejected",
                status="error",
                action=request.action,
                code=route_error["code"],
                error=route_error["error"],
                supportedApps=route_error.get("supportedApps"),
            )
            return self._complete(trace, started, request.action, route_error)

        selected_app = route["app"]
        trace.event(
            "route.selected",
            action=request.action,
            app=selected_app,
            source=route["source"],
            supportedApps=route["supportedApps"],
        )

        if sys.platform == "win32":
            try:
                self._catalog.validate_params(
                    selected_app,
                    request.action,
                    action_params,
                )
            except ActionValidationError as exc:
                result = {
                    "success": False,
                    "code": "INVALID_PARAMS",
                    "error": str(exc),
                }
                trace.event(
                    "dispatch.rejected",
                    status="error",
                    action=request.action,
                    app=selected_app,
                    code=result["code"],
                    error=result["error"],
                )
                trace.event(
                    "action.completed",
                    status="error",
                    action=request.action,
                    code=result["code"],
                    error=result["error"],
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )
                decorated = trace.decorate(result)
                return ActionResponse(decorated, trace.trace_id, trace.log_path)

        self._execution_lock.acquire()
        try:
            with self._state_condition:
                runtime_open = self._state == "open"
            if not runtime_open:
                result = {
                    "success": False,
                    "code": "RUNTIME_CLOSED",
                    "error": "Action Runtime 已关闭",
                }
                trace.event(
                    "action.completed",
                    status="error",
                    action=request.action,
                    code=result["code"],
                    error=result["error"],
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )
                decorated = trace.decorate(result)
                return ActionResponse(decorated, trace.trace_id, trace.log_path)
            if selected_app == "bridge":
                result = {"success": True, "data": {}}
                for app in ("excel", "ppt", "word"):
                    ping_started = time.perf_counter()
                    try:
                        with self._controller_lock:
                            controller = self._controllers.get(app)
                            if controller is None or not getattr(
                                controller,
                                "_ready",
                                False,
                            ):
                                controller = self._controller_factory(app, trace=trace)
                                self._controllers[app] = controller
                        available = bool(controller.ping(trace=trace))
                        result["data"][app] = available
                        trace.event(
                            "ping.app.completed",
                            app=app,
                            status="success" if available else "unavailable",
                            elapsedMs=round(
                                (time.perf_counter() - ping_started) * 1000,
                                2,
                            ),
                        )
                    except Exception as exc:
                        result["data"][app] = False
                        result["data"][f"{app}_error"] = str(exc)
                        trace.event(
                            "ping.app.completed",
                            app=app,
                            status="error",
                            error=f"{type(exc).__name__}: {exc}",
                            elapsedMs=round(
                                (time.perf_counter() - ping_started) * 1000,
                                2,
                            ),
                        )
                if sys.platform == "win32":
                    try:
                        self._catalog.validate_result(
                            "bridge",
                            request.action,
                            result["data"],
                        )
                    except ActionValidationError as exc:
                        result = {
                            "success": False,
                            "code": "INVALID_RESULT",
                            "error": str(exc),
                        }
                return self._complete(trace, started, request.action, result)
            controller_started = time.perf_counter()
            try:
                with self._controller_lock:
                    controller = self._controllers.get(selected_app)
                    if controller is None or not getattr(controller, "_ready", False):
                        controller = self._controller_factory(selected_app, trace=trace)
                        self._controllers[selected_app] = controller
            except Exception as exc:
                result = {
                    "success": False,
                    "code": "CONTROLLER_INIT_FAILED",
                    "error": f"{selected_app} 控制器初始化失败: {exc}",
                }
                trace.event(
                    "controller.init.failed",
                    status="error",
                    app=selected_app,
                    controller=f"Wps{selected_app.title()}Controller",
                    error=f"{type(exc).__name__}: {exc}",
                    elapsedMs=round(
                        (time.perf_counter() - controller_started) * 1000,
                        2,
                    ),
                )
                trace.event(
                    "action.completed",
                    status="error",
                    action=request.action,
                    code=result["code"],
                    error=result["error"],
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )
                decorated = trace.decorate(result)
                return ActionResponse(decorated, trace.trace_id, trace.log_path)

            platform_name = getattr(controller, "platform", "unknown")
            backend_kind = {
                "Windows": "powershell_com",
                "Linux": "openxml_file",
            }.get(platform_name, "unsupported")
            trace.event(
                "controller.ready",
                app=selected_app,
                controller=type(controller).__name__,
                platform=platform_name,
                backend=backend_kind,
                progId=_APP_PROGIDS[selected_app],
                elapsedMs=round((time.perf_counter() - controller_started) * 1000, 2),
            )

            execute_started = time.perf_counter()
            try:
                result = controller.execute(request.action, action_params, trace=trace)
                if not isinstance(result, dict):
                    result = {"success": True, "data": result}
            except Exception as exc:
                result = {"success": False, "error": f"执行失败: {exc}"}
            if (
                result.get("success")
                and sys.platform == "win32"
                and platform_name == "Windows"
            ):
                try:
                    self._catalog.validate_result(
                        selected_app,
                        request.action,
                        result.get("data", {}),
                    )
                except ActionValidationError as exc:
                    result = {
                        "success": False,
                        "code": "INVALID_RESULT",
                        "error": str(exc),
                    }
        finally:
            self._execution_lock.release()
        trace.event(
            "controller.execute.completed",
            status="success" if result.get("success") else "error",
            app=selected_app,
            action=request.action,
            code=result.get("code"),
            error=result.get("error"),
            elapsedMs=round((time.perf_counter() - execute_started) * 1000, 2),
            dispatchElapsedMs=round((time.perf_counter() - started) * 1000, 2),
            **trace.debug_fields(response=result),
        )
        trace.event(
            "action.completed",
            status="success" if result.get("success") else "error",
            action=request.action,
            code=result.get("code"),
            error=result.get("error"),
            elapsedMs=round((time.perf_counter() - started) * 1000, 2),
        )
        decorated = trace.decorate(result)
        return ActionResponse(decorated, trace.trace_id, trace.log_path)

    def close(self) -> None:
        with self._state_condition:
            if self._state == "closed":
                return
            if self._state == "closing":
                while self._state != "closed":
                    self._state_condition.wait()
                return
            self._state = "closing"
        try:
            with self._execution_lock:
                with self._controller_lock:
                    controllers = list(self._controllers.values())
                    self._controllers.clear()
                for controller in controllers:
                    try:
                        controller.close()
                    except Exception:
                        pass
        finally:
            with self._state_condition:
                self._state = "closed"
                self._state_condition.notify_all()
