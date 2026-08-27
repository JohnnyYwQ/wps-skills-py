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

    def test_concurrent_calls_are_serialized_and_reuse_one_ready_controller(self):
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
        self.assertEqual(["excel"], factory_calls)

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


if __name__ == "__main__":
    unittest.main()
