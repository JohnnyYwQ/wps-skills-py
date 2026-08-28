import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from io import StringIO
from unittest.mock import Mock, patch

from action_runtime import ActionRequest


SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))
import call


class CallRuntimeAdapterTests(unittest.TestCase):
    def test_params_file_and_app_are_parsed_without_mutating_action_params(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "params.json"
            path.write_text('{"find_text":"old"}', encoding="utf-8")

            action, params, app = call._load_params_from_args([
                "findReplace",
                "--app",
                "word",
                "--params-file",
                str(path),
            ])

        self.assertEqual("findReplace", action)
        self.assertEqual({"find_text": "old"}, params)
        self.assertEqual("word", app)

    def test_all_parameter_sources_reach_the_same_runtime_boundary(self):
        params = {"row": 1, "col": 1, "value": 42}
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "params.json"
            path.write_text(json.dumps(params), encoding="utf-8")
            variants = (
                ("inline", ["setCellValue", "--app", "excel", json.dumps(params)], ""),
                ("file", ["setCellValue", "--app", "excel", "--params-file", str(path)], ""),
                ("stdin", ["setCellValue", "--app", "excel", "--stdin"], json.dumps(params)),
            )
            for source, arguments, stdin in variants:
                with self.subTest(source=source):
                    runtime = Mock()
                    response = Mock()
                    response.to_dict.return_value = {"success": True, "data": {}}
                    runtime.execute.return_value = response
                    output = StringIO()

                    with patch.object(
                        call,
                        "ActionRuntime",
                        return_value=runtime,
                    ), patch.object(
                        call.sys,
                        "argv",
                        ["call.py", *arguments],
                    ), patch.object(
                        call.sys,
                        "stdin",
                        StringIO(stdin),
                    ), patch("sys.stdout", output), self.assertRaises(
                        SystemExit,
                    ) as exited:
                        call.main()

                self.assertEqual(0, exited.exception.code)
                runtime.execute.assert_called_once_with(ActionRequest(
                    action="setCellValue",
                    params=params,
                    app="excel",
                ))
                runtime.close.assert_called_once_with()
                self.assertEqual({"success": True, "data": {}}, json.loads(output.getvalue()))

    def test_cli_executes_in_process_and_always_closes_runtime(self):
        runtime = Mock()
        response = Mock()
        response.to_dict.return_value = {
            "success": True,
            "data": {"written": True},
            "traceId": "act-20260827-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
            "traceLog": "C:/tmp/action.jsonl",
        }
        runtime.execute.return_value = response

        with patch.object(
            call,
            "_load_params_from_args",
            return_value=("setCellValue", {"row": 1, "col": 1, "value": 42}, "excel"),
        ), patch.object(
            call,
            "ActionRuntime",
            return_value=runtime,
        ), patch("builtins.print") as output, self.assertRaises(
            SystemExit,
        ) as exited:
            call.main()

        self.assertEqual(0, exited.exception.code)
        runtime.execute.assert_called_once_with(ActionRequest(
            action="setCellValue",
            params={"row": 1, "col": 1, "value": 42},
            app="excel",
        ))
        runtime.close.assert_called_once_with()
        self.assertFalse(hasattr(call, "_ensure_server"))
        self.assertFalse(hasattr(call, "_post"))
        self.assertEqual(
            response.to_dict.return_value,
            json.loads(output.call_args.args[0]),
        )

    def test_cli_returns_exit_one_for_structured_runtime_error(self):
        runtime = Mock()
        response = Mock()
        response.to_dict.return_value = {
            "success": False,
            "code": "INVALID_PARAMS",
            "error": "params.slideIndex must be integer",
            "traceId": "act-20260827-bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
            "traceLog": "C:/tmp/action.jsonl",
        }
        runtime.execute.return_value = response

        with patch.object(
            call,
            "_load_params_from_args",
            return_value=("deleteSlide", {"slideIndex": "1"}, "ppt"),
        ), patch.object(
            call,
            "ActionRuntime",
            return_value=runtime,
        ), patch("builtins.print"), self.assertRaises(SystemExit) as exited:
            call.main()

        self.assertEqual(1, exited.exception.code)
        runtime.close.assert_called_once_with()

    def test_invalid_stdin_json_is_a_structured_failure(self):
        runtime = Mock()
        output = StringIO()

        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp, "WPS_TRACE": "info"},
            clear=False,
        ), patch.object(
            call,
            "ActionRuntime",
            return_value=runtime,
        ), patch.object(
            call.sys,
            "argv",
            ["call.py", "ping", "--stdin"],
        ), patch.object(
            call.sys,
            "stdin",
            StringIO("{not json}"),
        ), patch("sys.stdout", output), self.assertRaises(SystemExit) as exited:
            call.main()

        result = json.loads(output.getvalue())
        self.assertEqual(1, exited.exception.code)
        self.assertFalse(result["success"])
        self.assertEqual("INVALID_PARAMS", result["code"])
        self.assertIn("traceId", result)
        runtime.execute.assert_not_called()
        runtime.close.assert_called_once_with()

    def test_cli_converts_unexpected_execution_error_to_json_and_closes_runtime(self):
        runtime = Mock()
        runtime.execute.side_effect = RuntimeError("controller disconnected")
        output = StringIO()

        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp, "WPS_TRACE": "info"},
            clear=False,
        ), patch.object(
            call,
            "_load_params_from_args",
            return_value=("setCellValue", {"row": 1, "col": 1, "value": 42}, "excel"),
        ), patch.object(
            call,
            "ActionRuntime",
            return_value=runtime,
        ), patch("sys.stdout", output), self.assertRaises(SystemExit) as exited:
            call.main()

        result = json.loads(output.getvalue())
        self.assertEqual(1, exited.exception.code)
        self.assertEqual("ACTION_EXECUTION_FAILED", result["code"])
        self.assertIn("traceId", result)
        runtime.close.assert_called_once_with()

    def test_cli_converts_non_json_runtime_response_to_json_and_closes_runtime(self):
        runtime = Mock()
        response = Mock()
        response.to_dict.return_value = {
            "success": True,
            "data": {"bad": object()},
            "traceId": "act-20260827-cccccccccccccccccccccccccccccccc",
            "traceLog": "C:/tmp/action.jsonl",
        }
        runtime.execute.return_value = response
        output = StringIO()

        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {"WPS_TRACE_DIR": tmp, "WPS_TRACE": "info"},
            clear=False,
        ), patch.object(
            call,
            "_load_params_from_args",
            return_value=("setCellValue", {"row": 1, "col": 1, "value": 42}, "excel"),
        ), patch.object(
            call,
            "ActionRuntime",
            return_value=runtime,
        ), patch("sys.stdout", output), self.assertRaises(SystemExit) as exited:
            call.main()

        result = json.loads(output.getvalue())
        self.assertEqual(1, exited.exception.code)
        self.assertEqual("OUTPUT_SERIALIZATION_FAILED", result["code"])
        self.assertEqual(
            "act-20260827-cccccccccccccccccccccccccccccccc",
            result["traceId"],
        )
        runtime.close.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
