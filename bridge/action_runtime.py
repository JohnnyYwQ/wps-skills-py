"""In-process Action execution behind the ``execute`` / ``close`` seam."""

from dataclasses import dataclass, field
import inspect
import os
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
from action_gate import WindowsActionGate


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

ACTION_QUEUE_TIMEOUT_SECONDS = 600
ACTION_EXECUTION_TIMEOUT_SECONDS = 120


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


def _default_controller_factory(app, trace=None, deadline=None):
    return _APP_MODULES[app].get_controller(trace=trace, deadline=deadline)


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


def _supports_deadline(callback):
    try:
        signature = inspect.signature(callback)
    except (TypeError, ValueError):
        return True
    return (
        "deadline" in signature.parameters
        or any(
            parameter.kind == inspect.Parameter.VAR_KEYWORD
            for parameter in signature.parameters.values()
        )
    )


class ActionRuntime:
    """Resolve, execute, trace, cache, and close local WPS Actions."""

    def __init__(
        self,
        *,
        catalog=None,
        controller_factory: Optional[Callable[..., Any]] = None,
        action_gate_factory: Optional[Callable[[], Any]] = None,
        clock: Optional[Callable[[], float]] = None,
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
        self._action_gate_factory = action_gate_factory
        self._clock = clock or time.monotonic
        self._controllers = {}
        self._controller_lock = threading.Lock()
        self._execution_lock = threading.Lock()
        self._state_condition = threading.Condition()
        self._state = "open"

    def _new_controller(self, app, trace, deadline):
        """Create a controller, supplying the action deadline when supported."""
        factory = self._controller_factory
        if _supports_deadline(factory):
            return factory(app, trace=trace, deadline=deadline)
        return factory(app, trace=trace)

    @staticmethod
    def _call_controller(controller, action, params, trace, deadline):
        execute = controller.execute
        if _supports_deadline(execute):
            return execute(action, params, trace=trace, deadline=deadline)
        return execute(action, params, trace=trace)

    @staticmethod
    def _ping_controller(controller, trace, deadline):
        ping = controller.ping
        if _supports_deadline(ping):
            return ping(trace=trace, deadline=deadline)
        return ping(trace=trace)

    def _gate_for_action(self):
        if self._action_gate_factory is not None:
            return self._action_gate_factory()
        if sys.platform == "win32" and os.name == "nt":
            return WindowsActionGate()
        return None

    def _close_controllers(self):
        with self._controller_lock:
            controllers = list(self._controllers.values())
            self._controllers.clear()
        for controller in controllers:
            try:
                controller.close()
            except Exception:
                pass

    def _execution_timeout(self, deadline):
        return self._clock() >= deadline

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

        gate = self._gate_for_action()
        if gate is not None:
            try:
                acquired = gate.acquire(ACTION_QUEUE_TIMEOUT_SECONDS)
            except Exception as exc:
                try:
                    gate.close()
                except Exception:
                    pass
                result = {
                    "success": False,
                    "code": "ACTION_GATE_FAILED",
                    "error": f"Action 串行边界不可用: {exc}",
                }
                return self._complete(trace, started, request.action, result)
            if not acquired:
                try:
                    gate.close()
                except Exception:
                    pass
                result = {
                    "success": False,
                    "code": "ACTION_QUEUE_TIMEOUT",
                    "error": "等待其他 Action 完成超过 600 秒",
                }
                trace.event(
                    "action.gate.timeout",
                    status="timeout",
                    timeoutSeconds=ACTION_QUEUE_TIMEOUT_SECONDS,
                )
                return self._complete(trace, started, request.action, result)
            execution_deadline = self._clock() + ACTION_EXECUTION_TIMEOUT_SECONDS
            trace.event(
                "action.gate.acquired",
                timeoutSeconds=ACTION_QUEUE_TIMEOUT_SECONDS,
                executionBudgetSeconds=ACTION_EXECUTION_TIMEOUT_SECONDS,
            )
        else:
            execution_deadline = None

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
                execution_timed_out = False
                for app in ("excel", "ppt", "word"):
                    if (
                        execution_deadline is not None
                        and self._execution_timeout(execution_deadline)
                    ):
                        execution_timed_out = True
                        break
                    ping_started = time.perf_counter()
                    try:
                        with self._controller_lock:
                            controller = self._controllers.get(app)
                            if controller is None or not getattr(
                                controller,
                                "_ready",
                                False,
                            ):
                                controller = self._new_controller(
                                    app,
                                    trace,
                                    execution_deadline,
                                )
                                self._controllers[app] = controller
                        if (
                            execution_deadline is not None
                            and self._execution_timeout(execution_deadline)
                        ):
                            raise TimeoutError("Action 执行预算已耗尽")
                        available = bool(self._ping_controller(
                            controller,
                            trace,
                            execution_deadline,
                        ))
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
                    except TimeoutError as exc:
                        execution_timed_out = True
                        trace.event(
                            "ping.app.completed",
                            app=app,
                            status="timeout",
                            error=f"{type(exc).__name__}: {exc}",
                            elapsedMs=round(
                                (time.perf_counter() - ping_started) * 1000,
                                2,
                            ),
                        )
                        break
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
                if execution_timed_out:
                    result = {
                        "success": False,
                        "code": "ACTION_EXECUTION_TIMEOUT",
                        "error": "Action 执行超过 120 秒预算",
                    }
                elif sys.platform == "win32":
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
                if (
                    execution_deadline is not None
                    and self._execution_timeout(execution_deadline)
                ):
                    result = {
                        "success": False,
                        "code": "ACTION_EXECUTION_TIMEOUT",
                        "error": "Action 执行超过 120 秒预算",
                    }
                return self._complete(trace, started, request.action, result)
            controller_started = time.perf_counter()
            try:
                with self._controller_lock:
                    controller = self._controllers.get(selected_app)
                    if controller is None or not getattr(controller, "_ready", False):
                        if (
                            execution_deadline is not None
                            and self._execution_timeout(execution_deadline)
                        ):
                            raise TimeoutError("Action 执行预算已耗尽")
                        controller = self._new_controller(
                            selected_app,
                            trace,
                            execution_deadline,
                        )
                        self._controllers[selected_app] = controller
            except Exception as exc:
                if (
                    execution_deadline is not None
                    and (
                        isinstance(exc, TimeoutError)
                        or self._execution_timeout(execution_deadline)
                    )
                ):
                    result = {
                        "success": False,
                        "code": "ACTION_EXECUTION_TIMEOUT",
                        "error": "Action 执行超过 120 秒预算",
                    }
                else:
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
                if (
                    execution_deadline is not None
                    and self._execution_timeout(execution_deadline)
                ):
                    raise TimeoutError("Action 执行预算已耗尽")
                result = self._call_controller(
                    controller,
                    request.action,
                    action_params,
                    trace,
                    execution_deadline,
                )
                if not isinstance(result, dict):
                    result = {"success": True, "data": result}
            except Exception as exc:
                if isinstance(exc, TimeoutError):
                    result = {
                        "success": False,
                        "code": "ACTION_EXECUTION_TIMEOUT",
                        "error": "Action 执行超过 120 秒预算",
                    }
                else:
                    result = {"success": False, "error": f"执行失败: {exc}"}
            if (
                execution_deadline is not None
                and self._execution_timeout(execution_deadline)
            ):
                result = {
                    "success": False,
                    "code": "ACTION_EXECUTION_TIMEOUT",
                    "error": "Action 执行超过 120 秒预算",
                }
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
            try:
                self._close_controllers()
            finally:
                if gate is not None:
                    try:
                        gate.release()
                    finally:
                        gate.close()
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
