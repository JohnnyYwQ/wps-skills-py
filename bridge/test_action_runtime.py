import json
import os
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

import action_runtime
from action_catalog import ActionManifestError
from action_runtime import ActionRequest, ActionRuntime


class _RecordingController:
    platform = "Linux"

    def __init__(self):
        self._ready = True
        self.calls = []
        self.closed = False
        self.close_count = 0
        self.ping_count = 0

    def execute(self, action, params, trace=None):
        self.calls.append((action, params, trace.trace_id))
        return {"success": True, "data": {}}

    def close(self):
        self._ready = False
        self.closed = True
        self.close_count += 1

    def ping(self, trace=None):
        self.ping_count += 1
        return True


class _ScriptedController:
    platform = "Windows"

    def __init__(self, responses):
        self._ready = True
        self._responses = iter(responses)
        self.calls = []
        self.close_count = 0

    def execute(self, action, params, trace=None, deadline=None, correlation_id=None):
        self.calls.append({
            "action": action,
            "params": params,
            "trace_id": trace.trace_id,
            "deadline": deadline,
            "correlation_id": correlation_id,
        })
        return next(self._responses)

    def close(self):
        self._ready = False
        self.close_count += 1


class _FakeActionGate:
    def __init__(self, *, acquired=True, on_acquire=None):
        self.acquired = acquired
        self.on_acquire = on_acquire
        self.waits = []
        self.release_count = 0
        self.close_count = 0

    def acquire(self, timeout_seconds):
        self.waits.append(timeout_seconds)
        if self.on_acquire:
            self.on_acquire()
        return self.acquired

    def release(self):
        self.release_count += 1

    def close(self):
        self.close_count += 1


class _FakeClock:
    def __init__(self):
        self.now = 0.0

    def monotonic(self):
        return self.now

    def advance(self, seconds):
        self.now += seconds


class ActionRuntimeTests(unittest.TestCase):
    def test_execute_routes_unique_action_and_returns_action_trace(self):
        controller = _RecordingController()
        requested_apps = []

        def controller_factory(app, trace=None):
            requested_apps.append(app)
            return controller

        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp, "WPS_TRACE": "info"},
            clear=False,
        ):
            runtime = ActionRuntime(controller_factory=controller_factory)
            try:
                response = runtime.execute(ActionRequest(
                    action="setCellValue",
                    params={"row": 1, "col": 1, "value": 42},
                ))
                result = response.to_dict()
                trace_rows = [
                    json.loads(line)
                    for line in response.trace_log.read_text(encoding="utf-8").splitlines()
                ]
            finally:
                runtime.close()

        self.assertTrue(result["success"])
        self.assertEqual(["excel"], requested_apps)
        self.assertEqual("setCellValue", controller.calls[0][0])
        self.assertEqual({"row": 1, "col": 1, "value": 42}, controller.calls[0][1])
        self.assertEqual(result["traceId"], controller.calls[0][2])
        self.assertEqual(str(response.trace_log), result["traceLog"])
        self.assertEqual(
            ["action.started", "route.selected", "controller.ready",
             "controller.execute.completed", "action.completed"],
            [row["event"] for row in trace_rows],
        )

    def test_ambiguous_action_is_rejected_before_controller_initialization(self):
        initialized = []

        def controller_factory(app, trace=None):
            initialized.append(app)
            return _RecordingController()

        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp, "WPS_TRACE": "info"},
            clear=False,
        ):
            runtime = ActionRuntime(controller_factory=controller_factory)
            try:
                response = runtime.execute(ActionRequest(
                    action="findReplace",
                    params={"find": "old", "replace": "new"},
                ))
                result = response.to_dict()
                events = [
                    json.loads(line)["event"]
                    for line in response.trace_log.read_text(encoding="utf-8").splitlines()
                ]
            finally:
                runtime.close()

        self.assertFalse(result["success"])
        self.assertEqual("AMBIGUOUS_ACTION", result["code"])
        self.assertEqual(["excel", "word"], result["supportedApps"])
        self.assertEqual([], initialized)
        self.assertEqual(
            ["action.started", "route.rejected", "action.completed"],
            events,
        )

    def test_windows_invalid_params_are_rejected_before_controller_initialization(self):
        initialized = []

        def controller_factory(app, trace=None):
            initialized.append(app)
            return _RecordingController()

        runtime = ActionRuntime(controller_factory=controller_factory)
        try:
            with patch.object(action_runtime.sys, "platform", "win32"):
                result = runtime.execute(ActionRequest(
                    action="deleteSlide",
                    app="ppt",
                    params={"slideIndex": "1"},
                )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertEqual("INVALID_PARAMS", result["code"])
        self.assertIn("params.slideIndex must be integer", result["error"])
        self.assertEqual([], initialized)

    def test_read_action_rebuilds_the_controller_once_after_rpc_disconnect(self):
        failed = _ScriptedController([
            {"success": False, "error": "RPC server is unavailable (0x800706BA)"},
        ])
        recovered = _ScriptedController([
            {"success": True, "data": {"sheets": [], "count": 0}},
        ])
        controllers = iter([failed, recovered])
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None, deadline=None: next(controllers),
        )
        try:
            response = runtime.execute(ActionRequest(
                action="getSheetList",
                app="excel",
            ))
            result = response.to_dict()
            events = [
                json.loads(line)["event"]
                for line in response.trace_log.read_text(encoding="utf-8").splitlines()
            ]
        finally:
            runtime.close()

        self.assertTrue(result["success"])
        self.assertEqual(1, len(failed.calls))
        self.assertEqual(1, len(recovered.calls))
        self.assertEqual(failed.calls[0]["trace_id"], recovered.calls[0]["trace_id"])
        self.assertNotEqual(
            failed.calls[0]["correlation_id"],
            recovered.calls[0]["correlation_id"],
        )
        self.assertEqual(1, failed.close_count)
        self.assertEqual(1, recovered.close_count)
        self.assertIn("controller.retry.scheduled", events)
        self.assertIn("controller.retry.rebuilt", events)
        self.assertIn("controller.retry.completed", events)

    def test_read_action_does_not_retry_a_non_transient_wps_error(self):
        controller = _ScriptedController([
            {"success": False, "error": "工作表名称已存在"},
        ])
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None, deadline=None: controller,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="getSheetList",
                app="excel",
            )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertFalse(result["outcomeUnknown"])
        self.assertEqual(1, len(controller.calls))

    def test_excel_no_active_workbook_is_a_clear_known_outcome(self):
        controller = _ScriptedController([
            {
                "success": False,
                "code": "NO_ACTIVE_DOCUMENT",
                "error": "没有活动工作簿；请先创建或打开工作簿",
            },
        ])
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None, deadline=None: controller,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="getCellValue",
                app="excel",
                params={"row": 1, "col": 1},
            )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertEqual("NO_ACTIVE_DOCUMENT", result["code"])
        self.assertFalse(result["outcomeUnknown"])
        self.assertEqual(1, len(controller.calls))

    def test_read_action_returns_second_transient_failure_without_more_retries(self):
        failed = _ScriptedController([
            {"success": False, "error": "RPC server is unavailable (0x800706BA)"},
        ])
        retried = _ScriptedController([
            {"success": False, "error": "RPC server is unavailable (0x800706BA)"},
        ])
        controllers = iter([failed, retried])
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None, deadline=None: next(controllers),
        )
        try:
            result = runtime.execute(ActionRequest(
                action="getSheetList",
                app="excel",
            )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertFalse(result["outcomeUnknown"])
        self.assertEqual(1, len(failed.calls))
        self.assertEqual(1, len(retried.calls))

    def test_read_retry_uses_the_original_execution_deadline(self):
        clock = _FakeClock()
        gate = _FakeActionGate()
        failed = _ScriptedController([
            {"success": False, "error": "RPC server is unavailable (0x800706BA)"},
        ])
        retried = _ScriptedController([
            {"success": True, "data": {"sheets": [], "count": 0}},
        ])
        original_execute = failed.execute

        def fail_after_one_minute(*args, **kwargs):
            clock.advance(60)
            return original_execute(*args, **kwargs)

        original_retry_execute = retried.execute

        def finish_after_the_budget(*args, **kwargs):
            clock.advance(61)
            return original_retry_execute(*args, **kwargs)

        failed.execute = fail_after_one_minute
        retried.execute = finish_after_the_budget
        received_deadlines = []

        def controller_factory(app, trace=None, deadline=None):
            received_deadlines.append(deadline)
            return [failed, retried][len(received_deadlines) - 1]

        runtime = ActionRuntime(
            controller_factory=controller_factory,
            action_gate_factory=lambda: gate,
            clock=clock.monotonic,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="getSheetList",
                app="excel",
            )).to_dict()
        finally:
            runtime.close()

        self.assertEqual("ACTION_EXECUTION_TIMEOUT", result["code"])
        self.assertFalse(result["outcomeUnknown"])
        self.assertEqual([120, 120], received_deadlines)
        self.assertEqual(1, len(failed.calls))
        self.assertEqual(1, len(retried.calls))

    def test_write_action_reports_unknown_outcome_without_retry_after_rpc_disconnect(self):
        controller = _ScriptedController([
            {"success": False, "error": "RPC server is unavailable (0x800706BA)"},
        ])
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None, deadline=None: controller,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="setCellValue",
                app="excel",
                params={"row": 1, "col": 1, "value": 42},
            )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertTrue(result["outcomeUnknown"])
        self.assertEqual(1, len(controller.calls))

    def test_destructive_action_reports_unknown_outcome_without_retry_after_rpc_disconnect(self):
        controller = _ScriptedController([
            {"success": False, "error": "RPC server is unavailable (0x800706BA)"},
        ])
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None, deadline=None: controller,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="deleteSlide",
                app="ppt",
                params={"slideIndex": 1},
            )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertTrue(result["outcomeUnknown"])
        self.assertEqual(1, len(controller.calls))

    def test_write_timeout_reports_an_unknown_outcome_without_retry(self):
        controller = _ScriptedController([
            {"success": False, "code": "ACTION_EXECUTION_TIMEOUT", "error": "action 执行超时"},
        ])
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None, deadline=None: controller,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="setCellValue",
                app="excel",
                params={"row": 1, "col": 1, "value": 42},
            )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertTrue(result["outcomeUnknown"])
        self.assertEqual(1, len(controller.calls))

    def test_destructive_timeout_reports_an_unknown_outcome_without_retry(self):
        controller = _ScriptedController([
            {"success": False, "code": "ACTION_EXECUTION_TIMEOUT", "error": "action 执行超时"},
        ])
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None, deadline=None: controller,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="deleteSlide",
                app="ppt",
                params={"slideIndex": 1},
            )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertTrue(result["outcomeUnknown"])
        self.assertEqual(1, len(controller.calls))

    def test_validation_failure_has_a_known_outcome(self):
        runtime = ActionRuntime(controller_factory=lambda app, trace=None: (
            self.fail("controller should not be initialized")
        ))
        try:
            with patch.object(action_runtime.sys, "platform", "win32"):
                result = runtime.execute(ActionRequest(
                    action="deleteSlide",
                    app="ppt",
                    params={"slideIndex": "one"},
                )).to_dict()
        finally:
            runtime.close()

        self.assertEqual("INVALID_PARAMS", result["code"])
        self.assertFalse(result["outcomeUnknown"])

    def test_windows_invalid_success_result_is_rejected_after_controller_execution(self):
        controller = _RecordingController()
        controller.platform = "Windows"
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None: controller,
        )
        try:
            with patch.object(action_runtime.sys, "platform", "win32"):
                result = runtime.execute(ActionRequest(
                    action="addSlide",
                    app="ppt",
                    params={},
                )).to_dict()
        finally:
            runtime.close()

        self.assertEqual(1, len(controller.calls))
        self.assertFalse(result["success"])
        self.assertEqual("INVALID_RESULT", result["code"])
        self.assertIn("result.slideIndex is required", result["error"])

    def test_linux_preserves_existing_params_without_windows_schema_rejection(self):
        controller = _RecordingController()
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None: controller,
        )
        try:
            with patch.object(action_runtime.sys, "platform", "linux"):
                result = runtime.execute(ActionRequest(
                    action="setCellValue",
                    app="excel",
                    params={"cell": "A1", "value": 42},
                )).to_dict()
        finally:
            runtime.close()

        self.assertTrue(result["success"])
        self.assertEqual(
            {"cell": "A1", "value": 42},
            controller.calls[0][1],
        )

    def test_windows_unknown_parameter_is_rejected_before_controller(self):
        initialized = []
        runtime = ActionRuntime(controller_factory=lambda app, trace=None: (
            initialized.append(app) or _RecordingController()
        ))
        try:
            with patch.object(action_runtime.sys, "platform", "win32"):
                result = runtime.execute(ActionRequest(
                    action="addSlide",
                    app="ppt",
                    params={"layout": "blank", "template": "extra"},
                )).to_dict()
        finally:
            runtime.close()

        self.assertEqual("INVALID_PARAMS", result["code"])
        self.assertIn("params.template is not allowed", result["error"])
        self.assertEqual([], initialized)

    def test_null_params_are_rejected_before_controller_initialization(self):
        initialized = []
        runtime = ActionRuntime(controller_factory=lambda app, trace=None: (
            initialized.append(app) or _RecordingController()
        ))
        try:
            result = runtime.execute(ActionRequest(
                action="setCellValue",
                app="excel",
                params=None,
            )).to_dict()
        finally:
            runtime.close()

        self.assertEqual("INVALID_PARAMS", result["code"])
        self.assertEqual([], initialized)

    def test_concurrent_calls_are_serialized_and_close_each_action_controller(self):
        entered = threading.Event()
        overlap = threading.Event()
        release = threading.Event()
        active_lock = threading.Lock()
        controller = _RecordingController()
        controller.active = 0

        def blocking_execute(action, params, trace=None):
            with active_lock:
                controller.active += 1
                if controller.active > 1:
                    overlap.set()
                entered.set()
            try:
                release.wait(timeout=1)
                return {"success": True, "data": {}}
            finally:
                with active_lock:
                    controller.active -= 1

        controller.execute = blocking_execute
        factory_calls = []

        def controller_factory(app, trace=None):
            factory_calls.append(app)
            return controller

        runtime = ActionRuntime(controller_factory=controller_factory)
        responses = []

        def call(value):
            responses.append(runtime.execute(ActionRequest(
                action="setCellValue",
                params={"row": 1, "col": 1, "value": value},
            )))

        first = threading.Thread(target=call, args=(1,))
        second = threading.Thread(target=call, args=(2,))
        try:
            first.start()
            self.assertTrue(entered.wait(timeout=1))
            second.start()
            time.sleep(0.05)
            overlapped_before_release = overlap.is_set()
            release.set()
            first.join(timeout=2)
            second.join(timeout=2)
        finally:
            release.set()
            runtime.close()

        self.assertFalse(overlapped_before_release)
        self.assertFalse(first.is_alive())
        self.assertFalse(second.is_alive())
        self.assertEqual(2, len(responses))
        self.assertEqual(["excel", "excel"], factory_calls)

    def test_close_is_idempotent_and_closed_runtime_rejects_new_actions(self):
        controller = _RecordingController()
        factory_calls = []

        def controller_factory(app, trace=None):
            factory_calls.append(app)
            return controller

        runtime = ActionRuntime(controller_factory=controller_factory)
        first = runtime.execute(ActionRequest(
            action="setCellValue",
            params={"row": 1, "col": 1, "value": 1},
        ))

        runtime.close()
        runtime.close()
        rejected = runtime.execute(ActionRequest(
            action="setCellValue",
            params={"row": 1, "col": 1, "value": 2},
        )).to_dict()

        self.assertTrue(first.success)
        self.assertEqual(1, controller.close_count)
        self.assertEqual(["excel"], factory_calls)
        self.assertFalse(rejected["success"])
        self.assertEqual("RUNTIME_CLOSED", rejected["code"])
        self.assertIn("traceId", rejected)

    def test_controller_initialization_failure_is_a_structured_traced_response(self):
        def controller_factory(app, trace=None):
            raise RuntimeError("WPS unavailable")

        runtime = ActionRuntime(controller_factory=controller_factory)
        try:
            response = runtime.execute(ActionRequest(
                action="setCellValue",
                params={"row": 1, "col": 1, "value": 1},
            ))
            result = response.to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertEqual("CONTROLLER_INIT_FAILED", result["code"])
        self.assertIn("WPS unavailable", result["error"])
        self.assertIn("traceId", result)

    def test_runtime_routes_only_unique_actions_without_an_explicit_app(self):
        controllers = {}

        def controller_factory(app, trace=None):
            controller = controllers.setdefault(app, _RecordingController())
            return controller

        runtime = ActionRuntime(controller_factory=controller_factory)
        try:
            unknown = runtime.execute(ActionRequest(
                action="notARealAction",
            )).to_dict()
            invalid_app = runtime.execute(ActionRequest(
                action="setCellValue",
                app="pages",
                params={"row": 1, "col": 1, "value": 1},
            )).to_dict()
            unsupported = runtime.execute(ActionRequest(
                action="setCellValue",
                app="word",
                params={"row": 1, "col": 1, "value": 1},
            )).to_dict()
            conflicting = runtime.execute(ActionRequest(
                action="findReplace",
                app="excel",
                params={"app": "word", "find": "a", "replace": "b"},
            )).to_dict()
            params_app = runtime.execute(ActionRequest(
                action="findReplace",
                params={"app": "word", "find_text": "a", "replace_text": "b"},
            )).to_dict()
            ambiguous_save = runtime.execute(ActionRequest(
                action="saveAs",
                params={"filePath": "C:/tmp/report.pptx"},
            )).to_dict()
        finally:
            runtime.close()

        self.assertEqual("UNKNOWN_ACTION", unknown["code"])
        self.assertEqual("INVALID_APP", invalid_app["code"])
        self.assertEqual("ACTION_NOT_SUPPORTED_FOR_APP", unsupported["code"])
        self.assertEqual(["excel"], unsupported["supportedApps"])
        self.assertEqual("CONFLICTING_APP", conflicting["code"])
        self.assertTrue(params_app["success"])
        self.assertNotIn("app", controllers["word"].calls[-1][1])
        self.assertEqual("AMBIGUOUS_ACTION", ambiguous_save["code"])
        self.assertEqual(
            ["excel", "ppt", "word"],
            ambiguous_save["supportedApps"],
        )
        self.assertNotIn("ppt", controllers)

    def test_bridge_ping_contract_aggregates_cached_application_controllers(self):
        controllers = {}
        factory_calls = []

        def controller_factory(app, trace=None):
            factory_calls.append(app)
            return controllers.setdefault(app, _RecordingController())

        runtime = ActionRuntime(controller_factory=controller_factory)
        try:
            result = runtime.execute(ActionRequest(action="ping")).to_dict()
        finally:
            runtime.close()

        self.assertTrue(result["success"])
        self.assertEqual(
            {"excel": True, "ppt": True, "word": True},
            result["data"],
        )
        self.assertEqual(["excel", "ppt", "word"], factory_calls)
        self.assertEqual(
            {"excel": 1, "ppt": 1, "word": 1},
            {app: controller.ping_count for app, controller in controllers.items()},
        )
        self.assertTrue(all(controller.closed for controller in controllers.values()))

    def test_bridge_action_rejects_an_explicit_application_before_controller_init(self):
        initialized = []
        runtime = ActionRuntime(controller_factory=lambda app, trace=None: (
            initialized.append(app) or _RecordingController()
        ))
        try:
            result = runtime.execute(ActionRequest(
                action="ping",
                app="excel",
            )).to_dict()
        finally:
            runtime.close()

        self.assertEqual("ACTION_NOT_SUPPORTED_FOR_APP", result["code"])
        self.assertEqual(["bridge"], result["supportedApps"])
        self.assertEqual([], initialized)

    def test_invalid_manifest_is_returned_as_stable_runtime_error(self):
        with patch.object(
            action_runtime.ActionCatalog,
            "from_path",
            side_effect=ActionManifestError("manifest is broken"),
        ):
            runtime = ActionRuntime(
                controller_factory=lambda app, trace=None: _RecordingController(),
            )
        try:
            result = runtime.execute(ActionRequest(
                action="setCellValue",
                params={"row": 1, "col": 1, "value": 1},
            )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertEqual("INVALID_ACTION_MANIFEST", result["code"])
        self.assertIn("manifest is broken", result["error"])
        self.assertIn("traceId", result)

    def test_action_queue_timeout_skips_controller_initialization_and_cleans_gate(self):
        gate = _FakeActionGate(acquired=False)
        initialized = []
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None: (
                initialized.append(app) or _RecordingController()
            ),
            action_gate_factory=lambda: gate,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="setCellValue",
                params={"row": 1, "col": 1, "value": 42},
            )).to_dict()
        finally:
            runtime.close()

        self.assertEqual("ACTION_QUEUE_TIMEOUT", result["code"])
        self.assertEqual([600], gate.waits)
        self.assertEqual([], initialized)
        self.assertEqual(0, gate.release_count)
        self.assertEqual(1, gate.close_count)

    def test_execution_budget_starts_after_the_gate_is_acquired(self):
        clock = _FakeClock()
        gate = _FakeActionGate(on_acquire=lambda: clock.advance(600))
        controller = _RecordingController()
        received_deadlines = []

        def controller_factory(app, trace=None, deadline=None):
            received_deadlines.append(deadline)
            return controller

        runtime = ActionRuntime(
            controller_factory=controller_factory,
            action_gate_factory=lambda: gate,
            clock=clock.monotonic,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="setCellValue",
                params={"row": 1, "col": 1, "value": 42},
            )).to_dict()
        finally:
            runtime.close()

        self.assertTrue(result["success"])
        self.assertEqual([720], received_deadlines)
        self.assertEqual(1, controller.close_count)
        self.assertEqual(1, gate.release_count)
        self.assertEqual(1, gate.close_count)

    def test_execution_timeout_closes_controller_before_releasing_gate(self):
        clock = _FakeClock()
        gate = _FakeActionGate()
        controller = _RecordingController()
        events = []

        def execute(action, params, trace=None, deadline=None):
            events.append("execute")
            clock.advance(121)
            return {"success": True, "data": {}}

        controller.execute = execute
        original_close = controller.close

        def close():
            events.append("close")
            original_close()

        controller.close = close
        original_release = gate.release

        def release():
            events.append("release")
            original_release()

        gate.release = release
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None, deadline=None: controller,
            action_gate_factory=lambda: gate,
            clock=clock.monotonic,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="setCellValue",
                params={"row": 1, "col": 1, "value": 42},
            )).to_dict()
        finally:
            runtime.close()

        self.assertEqual("ACTION_EXECUTION_TIMEOUT", result["code"])
        self.assertEqual(["execute", "close", "release"], events)

    def test_controller_initialization_cannot_consume_the_execution_budget(self):
        clock = _FakeClock()
        gate = _FakeActionGate()
        controller = _RecordingController()

        def controller_factory(app, trace=None, deadline=None):
            clock.advance(121)
            return controller

        runtime = ActionRuntime(
            controller_factory=controller_factory,
            action_gate_factory=lambda: gate,
            clock=clock.monotonic,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="setCellValue",
                params={"row": 1, "col": 1, "value": 42},
            )).to_dict()
        finally:
            runtime.close()

        self.assertEqual("ACTION_EXECUTION_TIMEOUT", result["code"])
        self.assertEqual([], controller.calls)
        self.assertEqual(1, controller.close_count)
        self.assertEqual(1, gate.release_count)

    def test_bridge_ping_stops_before_initializing_another_app_after_timeout(self):
        clock = _FakeClock()
        gate = _FakeActionGate()
        initialized = []

        def controller_factory(app, trace=None, deadline=None):
            controller = _RecordingController()
            initialized.append(app)
            if app == "excel":
                controller.ping = lambda trace=None: (clock.advance(121) or False)
            return controller

        runtime = ActionRuntime(
            controller_factory=controller_factory,
            action_gate_factory=lambda: gate,
            clock=clock.monotonic,
        )
        try:
            result = runtime.execute(ActionRequest(action="ping")).to_dict()
        finally:
            runtime.close()

        self.assertEqual("ACTION_EXECUTION_TIMEOUT", result["code"])
        self.assertEqual(["excel"], initialized)
        self.assertEqual(1, gate.release_count)

    def test_action_exception_releases_gate_after_controller_cleanup(self):
        gate = _FakeActionGate()
        controller = _RecordingController()
        controller.execute = lambda action, params, trace=None: (_ for _ in ()).throw(
            RuntimeError("controller disconnected")
        )
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None: controller,
            action_gate_factory=lambda: gate,
        )
        try:
            result = runtime.execute(ActionRequest(
                action="setCellValue",
                params={"row": 1, "col": 1, "value": 42},
            )).to_dict()
        finally:
            runtime.close()

        self.assertFalse(result["success"])
        self.assertEqual(1, controller.close_count)
        self.assertEqual(1, gate.release_count)
        self.assertEqual(1, gate.close_count)

    def test_interrupt_during_cleanup_still_releases_and_closes_gate(self):
        gate = _FakeActionGate()
        controller = _RecordingController()
        controller.close = lambda: (_ for _ in ()).throw(KeyboardInterrupt())
        runtime = ActionRuntime(
            controller_factory=lambda app, trace=None: controller,
            action_gate_factory=lambda: gate,
        )
        try:
            with self.assertRaises(KeyboardInterrupt):
                runtime.execute(ActionRequest(
                    action="setCellValue",
                    params={"row": 1, "col": 1, "value": 42},
                ))
        finally:
            runtime.close()

        self.assertEqual(1, gate.release_count)
        self.assertEqual(1, gate.close_count)


if __name__ == "__main__":
    unittest.main()
