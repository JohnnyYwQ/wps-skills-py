import json
import re
import tempfile
import unittest
from pathlib import Path

from action_catalog import (
    ActionCatalog,
    ActionManifestError,
    ActionValidationError,
    validate_windows_implementation_consistency,
    load_action_manifest,
    validate_instance,
)
import wps_excel
import wps_ppt
import wps_word


class ActionManifestValidationTests(unittest.TestCase):
    def _write_manifest(self, manifest, directory):
        path = Path(directory) / "action_manifest.json"
        path.write_text(json.dumps(manifest), encoding="utf-8")
        return path

    def test_unknown_schema_keyword_is_rejected_with_its_path(self):
        manifest = {
            "schema_version": 1,
            "actions": [
                {
                    "owner": "ppt",
                    "action": "addSlide",
                    "description": "新增幻灯片。",
                    "parameters": {
                        "type": "object",
                        "properties": {},
                        "required": [],
                        "additionalProperties": False,
                        "default": {},
                    },
                    "result": {
                        "type": "object",
                        "properties": {},
                        "required": [],
                        "additionalProperties": False,
                    },
                    "prerequisites": ["active_presentation"],
                    "risk": "write",
                }
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "action_manifest.json"
            path.write_text(json.dumps(manifest), encoding="utf-8")

            with self.assertRaises(ActionManifestError) as raised:
                load_action_manifest(path)

        self.assertEqual("INVALID_ACTION_MANIFEST", raised.exception.code)
        self.assertIn("actions[0].parameters.default", str(raised.exception))

    def test_manifest_structure_and_contract_identity_are_validated(self):
        empty_schema = {
            "type": "object",
            "properties": {},
            "required": [],
            "additionalProperties": False,
        }
        contract = {
            "owner": "ppt",
            "action": "addSlide",
            "description": "新增幻灯片。",
            "parameters": empty_schema,
            "result": empty_schema,
            "prerequisites": ["active_presentation"],
            "risk": "write",
        }
        invalid_manifests = (
            ({"schema_version": True, "actions": []}, "schema_version"),
            ({"schema_version": 2, "actions": []}, "schema_version"),
            (
                {"schema_version": 1, "actions": [contract, dict(contract)]},
                "duplicate contract",
            ),
            (
                {
                    "schema_version": 1,
                    "actions": [dict(contract, owner="spreadsheet")],
                },
                "actions[0].owner",
            ),
            (
                {
                    "schema_version": 1,
                    "actions": [dict(contract, owner=["ppt"])],
                },
                "actions[0].owner",
            ),
            (
                {
                    "schema_version": 1,
                    "actions": [dict(contract, risk=["write"])],
                },
                "actions[0].risk",
            ),
            (
                {
                    "schema_version": 1,
                    "actions": [{
                        **contract,
                        "parameters": {
                            "type": ["object"],
                            "properties": {},
                            "required": [],
                            "additionalProperties": False,
                        },
                    }],
                },
                "parameters.type",
            ),
            (
                {
                    "schema_version": 1,
                    "routing_defaults": {"addSlide": "ppt"},
                    "actions": [contract],
                },
                "routing_defaults",
            ),
            (
                {
                    "schema_version": 1,
                    "actions": [{
                        **contract,
                        "parameters": {
                            "type": "object",
                            "properties": {"layout": {"type": "string", "enum": [1]}},
                            "required": [],
                            "additionalProperties": False,
                        },
                    }],
                },
                "parameters.properties.layout.enum[0]",
            ),
        )
        with tempfile.TemporaryDirectory() as directory:
            for manifest, expected in invalid_manifests:
                with self.subTest(expected=expected):
                    path = self._write_manifest(manifest, directory)
                    with self.assertRaises(ActionManifestError) as raised:
                        load_action_manifest(path)
                    self.assertIn(expected, str(raised.exception))

    def test_supported_schema_subset_validates_values_with_field_paths(self):
        schema = {
            "type": "object",
            "properties": {
                "slideIndex": {"type": "integer", "minimum": 1},
                "layout": {"type": "string", "enum": ["blank", "title"]},
                "label": {"type": "string", "pattern": "^[A-Z]+$"},
                "value": {
                    "anyOf": [
                        {"type": "number"},
                        {"type": "string"},
                        {"type": "null"},
                    ]
                },
            },
            "required": ["slideIndex"],
            "additionalProperties": False,
        }
        validate_instance(
            {"slideIndex": 1, "layout": "blank", "label": "OK", "value": None},
            schema,
            "params",
        )

        invalid_values = (
            ({}, "params.slideIndex is required"),
            ({"slideIndex": True}, "params.slideIndex must be integer"),
            ({"slideIndex": 0}, "params.slideIndex must be >= 1"),
            ({"slideIndex": 1, "layout": "notes"}, "params.layout must be one of"),
            ({"slideIndex": 1, "label": "no"}, "params.label must match"),
            ({"slideIndex": 1, "extra": 2}, "params.extra is not allowed"),
            ({"slideIndex": 1, "value": False}, "params.value must match anyOf"),
        )
        for value, expected in invalid_values:
            with self.subTest(value=value):
                with self.assertRaises(ActionValidationError) as raised:
                    validate_instance(value, schema, "params")
                self.assertIn(expected, str(raised.exception))

    def test_contract_examples_are_checked_against_parameter_and_result_schemas(self):
        integer_object = {
            "type": "object",
            "properties": {"slideIndex": {"type": "integer"}},
            "required": ["slideIndex"],
            "additionalProperties": False,
        }
        contract = {
            "owner": "ppt",
            "action": "addSlide",
            "description": "新增幻灯片。",
            "parameters": integer_object,
            "result": integer_object,
            "prerequisites": [],
            "risk": "write",
            "examples": [{"params": {"slideIndex": "one"}, "result": {"slideIndex": 1}}],
        }
        with tempfile.TemporaryDirectory() as directory:
            path = self._write_manifest(
                {"schema_version": 1, "actions": [contract]}, directory
            )
            with self.assertRaises(ActionManifestError) as raised:
                load_action_manifest(path)

        self.assertIn("actions[0].examples[0].params.slideIndex", str(raised.exception))

    def test_default_manifest_matches_public_windows_handlers_in_both_directions(self):
        catalog = ActionCatalog.from_path()

        validate_windows_implementation_consistency(
            catalog,
            {
                "excel": wps_excel.PS_BRIDGE_SCRIPT,
                "ppt": wps_ppt.PS_BRIDGE_SCRIPT,
                "word": wps_word.PS_BRIDGE_SCRIPT,
            },
        )

        self.assertEqual(235, len(catalog.list()))
        self.assertEqual({"ping", "wireCheck"}, {
            item["action"] for item in catalog.list(owner="bridge")
        })

    def test_shared_chart_type_contract_cannot_drift_between_controllers(self):
        catalog = ActionCatalog.from_path()
        drifted_ppt = wps_ppt.PS_BRIDGE_SCRIPT.replace(
            '"scatter" { return -4169 }',
            '"bubble" { return 15 }',
            1,
        )

        with self.assertRaises(ActionManifestError) as raised:
            validate_windows_implementation_consistency(
                catalog,
                {
                    "excel": wps_excel.PS_BRIDGE_SCRIPT,
                    "ppt": drifted_ppt,
                    "word": wps_word.PS_BRIDGE_SCRIPT,
                },
            )

        self.assertIn("implementation.ppt.chartType", str(raised.exception))

    def test_excel_contracts_expose_explicit_safe_overwrite(self):
        catalog = ActionCatalog.from_path()

        for action, path_field in (
            ("saveAs", "filePath"),
            ("convertToPDF", "outputPath"),
            ("convertFormat", "outputPath"),
        ):
            with self.subTest(action=action):
                contract = catalog.get("excel", action)
                self.assertEqual("destructive", contract["risk"])
                self.assertEqual(
                    "boolean", contract["parameters"]["properties"]["overwrite"]["type"],
                )
                params = {"targetFormat": "xlsx"} if action == "convertFormat" else {}
                if action == "saveAs":
                    params[path_field] = r"C:\\tmp\\report.xlsx"
                catalog.validate_params("excel", action, {**params, "overwrite": True})
                with self.assertRaises(ActionValidationError):
                    catalog.validate_params("excel", action, {**params, "overwrite": "yes"})

    def test_excel_powershell_contract_preserves_active_workbook_and_targets(self):
        script = wps_excel.PS_BRIDGE_SCRIPT
        catalog = ActionCatalog.from_path()

        self.assertLess(script.index("GetActiveObject('Ket.Application')"), script.index("New-Object -ComObject 'Ket.Application'"))
        startup = script[:script.index("function Exec-ping")]
        self.assertNotIn("Workbooks.Add", startup)
        self.assertIn("NO_ACTIVE_DOCUMENT", script)
        allowed_actions = re.search(
            r"\$global:ExcelActionsWithoutActiveWorkbook = @\((.*?)\)",
            script,
            re.DOTALL,
        )
        self.assertIsNotNone(allowed_actions)
        self.assertEqual(
            {"ping"} | {
                summary["action"]
                for summary in catalog.list(owner="excel")
                if "active_workbook" not in catalog.get(
                    "excel", summary["action"],
                )["prerequisites"]
            },
            set(re.findall(r'"([A-Za-z][A-Za-z0-9]*)"', allowed_actions.group(1))),
        )
        self.assertRegex(
            script,
            r'if \(\(Test-ExcelActionRequiresActiveWorkbook \$action\) -and -not \(Get-ExcelActiveWorkbook\)\) \{\s*\$result = @\{\s*success=\$false\s*code="NO_ACTIVE_DOCUMENT"',
        )
        self.assertRegex(
            script,
            re.compile(r"function Exec-createWorkbook\(\$p\).*Workbooks\.Add", re.DOTALL),
        )
        self.assertRegex(
            script,
            re.compile(r"function Exec-openWorkbook\(\$p\).*Workbooks\.Open", re.DOTALL),
        )
        self.assertNotRegex(script, r"\.Quit\s*\(")
        self.assertNotIn("Remove-Item", script)
        self.assertIn("function Invoke-ExcelSafeTargetWrite", script)
        self.assertEqual(4, script.count("Invoke-ExcelSafeTargetWrite"))
        self.assertIn("TARGET_EXISTS", script)
        self.assertIn("OVERWRITE_NOT_SAFE", script)
        self.assertIn("OVERWRITE_RESTORE_FAILED", script)
        self.assertIn("[System.IO.File]::Copy($backupPath, $fullPath, $true)", script)
        self.assertEqual(120, wps_excel.EXEC_TIMEOUT)


if __name__ == "__main__":
    unittest.main()
