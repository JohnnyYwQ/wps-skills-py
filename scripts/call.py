#!/usr/bin/env python3
"""Thin command-line adapter for the in-process WPS Action Runtime."""

import json
import os
import sys


BRIDGE_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "bridge",
)
if BRIDGE_DIR not in sys.path:
    sys.path.insert(0, BRIDGE_DIR)

from action_runtime import ActionRequest, ActionRuntime
from action_trace import ActionTrace


class _InputError(ValueError):
    """The CLI could not turn its input into one Action request."""


def _usage_error(message):
    raise _InputError(message)


def _failure_response(code, error, trace_id=None):
    trace = ActionTrace.resume(trace_id, component="call")
    trace.event("cli.failed", status="error", code=code, error=error)
    return trace.decorate({"success": False, "code": code, "error": error})


def _trace_id_from(result):
    if isinstance(result, dict):
        return result.get("traceId")
    return None


def _serialize_response(result):
    try:
        return json.dumps(result, ensure_ascii=False, allow_nan=False), bool(
            isinstance(result, dict) and result.get("success")
        )
    except (TypeError, ValueError, OverflowError) as exc:
        fallback = _failure_response(
            "OUTPUT_SERIALIZATION_FAILED",
            f"Action 响应无法序列化为 JSON: {type(exc).__name__}: {exc}",
            _trace_id_from(result),
        )
        return json.dumps(fallback, ensure_ascii=False, allow_nan=False), False


def _load_params_from_args(args):
    """Return ``(action, params, app)`` for JSON, file, or stdin input."""
    action = None
    params = None
    params_source = None
    app = None
    use_stdin = False

    def set_params(value, source):
        nonlocal params, params_source
        if params_source is not None:
            _usage_error(
                f"参数只能提供一次，不能同时使用 {params_source} 和 {source}"
            )
        params = value
        params_source = source

    def load_json(raw, source):
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            _usage_error(f"{source} 内容不是合法 JSON: {exc}")

    index = 0
    while index < len(args):
        argument = args[index]
        if argument == "--stdin":
            use_stdin = True
            index += 1
            continue
        if argument == "--app":
            index += 1
            if index >= len(args):
                _usage_error("--app 缺少应用参数（excel/ppt/word）")
            new_app = args[index].strip().lower()
            if app is not None and app != new_app:
                _usage_error(f"重复的 --app 参数不一致: {app} / {new_app}")
            app = new_app
            index += 1
            continue
        if argument == "--params-file":
            index += 1
            if index >= len(args):
                _usage_error("--params-file 缺少文件路径参数")
            path = args[index]
            try:
                with open(path, "r", encoding="utf-8") as stream:
                    set_params(json.load(stream), "--params-file")
            except FileNotFoundError:
                _usage_error(f"--params-file 指定的文件不存在: {path}")
            except json.JSONDecodeError as exc:
                _usage_error(f"--params-file 内容不是合法 JSON: {exc}")
            except Exception as exc:
                _usage_error(f"读取 --params-file 失败: {exc}")
            index += 1
            continue
        if action is None:
            action = argument
        elif params_source is None:
            set_params(load_json(argument, "params"), "内联 JSON")
        else:
            _usage_error(f"不支持的额外参数: {argument}")
        index += 1

    if use_stdin and params_source is not None:
        _usage_error("--stdin 不能与内联 JSON 或 --params-file 同时使用")
    if params_source is None and (use_stdin or not sys.stdin.isatty()):
        raw = sys.stdin.read()
        if raw.strip():
            set_params(load_json(raw, "标准输入"), "标准输入")

    if action is None:
        _usage_error(
            "用法: call.py <action> [--app excel|ppt|word] "
            "['<json>' | --params-file <path> | --stdin]"
        )
    if params_source is None:
        params = {}
    return action, params, app


def main():
    runtime = None
    result = None
    try:
        runtime = ActionRuntime()
        action, params, app = _load_params_from_args(sys.argv[1:])
        response = runtime.execute(ActionRequest(
            action=action,
            params=params,
            app=app,
        ))
        result = response.to_dict()
    except _InputError as exc:
        result = _failure_response("INVALID_PARAMS", str(exc))
    except Exception as exc:
        result = _failure_response(
            "ACTION_EXECUTION_FAILED",
            f"Action CLI 执行失败: {type(exc).__name__}: {exc}",
        )
    finally:
        if runtime is not None:
            try:
                runtime.close()
            except Exception as exc:
                result = _failure_response(
                    "RUNTIME_CLOSE_FAILED",
                    f"Action Runtime 清理失败: {type(exc).__name__}: {exc}",
                    _trace_id_from(result),
                )

    encoded, success = _serialize_response(result)
    print(encoded)
    raise SystemExit(0 if success else 1)


if __name__ == "__main__":
    main()
