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
        self.assertNotIn("ExcelActionsWithoutActiveWorkbook", script)
        self.assertRegex(
            script,
            r'if \(\$cmd\.requiresActiveWorkbook -and -not \(Get-ExcelActiveWorkbook\)\) \{\s*\$result = @\{\s*success=\$false\s*code="NO_ACTIVE_DOCUMENT"',
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

    def test_excel_copy_sheet_keeps_the_copy_in_the_active_workbook(self):
        script = wps_excel.PS_BRIDGE_SCRIPT

        self.assertNotRegex(
            script,
            re.compile(
                r"function Exec-copySheet\(\$p\).*?\$sheet\.Copy\(\)",
                re.DOTALL,
            ),
        )

    def test_excel_action_response_ignores_com_pipeline_output(self):
        script = wps_excel.PS_BRIDGE_SCRIPT

        self.assertNotIn('$result = & "Exec-$action" $params', script)
        self.assertRegex(
            script,
            re.compile(
                r'\$actionOutput = @\(& "Exec-\$action" \$params\)\s*'
                r'\$result = \$actionOutput\[-1\]',
                re.DOTALL,
            ),
        )

    def test_ppt_and_word_action_responses_ignore_com_pipeline_output(self):
        for owner, script in (
            ("ppt", wps_ppt.PS_BRIDGE_SCRIPT),
            ("word", wps_word.PS_BRIDGE_SCRIPT),
        ):
            with self.subTest(owner=owner):
                self.assertNotIn('$result = & "Exec-$action" $params', script)
                self.assertRegex(
                    script,
                    re.compile(
                        r'\$actionOutput = @\(& "Exec-\$action" \$params\)\s*'
                        r'\$result = \$actionOutput\[-1\]',
                        re.DOTALL,
                    ),
                )

    def test_excel_text_actions_use_valid_powershell_and_range_properties(self):
        script = wps_excel.PS_BRIDGE_SCRIPT

        self.assertNotIn('if ($p.matchCase) { true }', script)
        self.assertIn('if ($p.matchCase) { $true }', script)
        self.assertNotIn('$global:excel.Selection.Text = $p.text', script)
        self.assertIn('$global:excel.Selection.Value2 = $p.text', script)

    def test_excel_formatting_uses_wps_available_color_and_window_paths(self):
        script = wps_excel.PS_BRIDGE_SCRIPT

        self.assertIn('function Convert-HexToOle($hex)', script)
        self.assertNotIn('[System.Drawing.ColorTranslator]', script)
        self.assertNotIn('$wb.ActiveWindow.', script)
        self.assertIn('$global:excel.ActiveWindow.FreezePanes', script)
        self.assertIn('$global:excel.ActiveWindow.Zoom', script)

    def test_excel_style_and_conditional_format_use_complete_com_arguments(self):
        script = wps_excel.PS_BRIDGE_SCRIPT

        self.assertIn('$style = $wb.Styles.Item($p.style)', script)
        self.assertIn('$sheet.Range($p.range).set_Style($style)', script)
        self.assertIn(
            '[void]$range.FormatConditions.Add(1, 5, $p.condition, [Type]::Missing)',
            script,
        )
        self.assertRegex(
            script,
            re.compile(
                r"function Exec-copySheet\(\$p\).*?"
                r"\[void\]\$sheet\.Copy\(\[Type\]::Missing, "
                r"\$wb\.Sheets\.Item\(\$wb\.Sheets\.Count\)\).*?"
                r"\$newSheet = \$wb\.Sheets\.Item\(\$wb\.Sheets\.Count\)",
                re.DOTALL,
            ),
        )

    def test_ppt_contracts_preserve_active_presentation_and_targets(self):
        catalog = ActionCatalog.from_path()
        script = wps_ppt.PS_BRIDGE_SCRIPT

        for action, path_field in (
            ("saveAs", "filePath"),
            ("convertToPDF", "outputPath"),
            ("convertFormat", "outputPath"),
        ):
            with self.subTest(action=action):
                contract = catalog.get("ppt", action)
                self.assertEqual("destructive", contract["risk"])
                self.assertEqual(
                    "boolean", contract["parameters"]["properties"]["overwrite"]["type"],
                )
                params = {"targetFormat": "pptx"} if action == "convertFormat" else {}
                if action == "saveAs":
                    params[path_field] = r"C:\\tmp\\report.pptx"
                catalog.validate_params("ppt", action, {**params, "overwrite": True})
                with self.assertRaises(ActionValidationError):
                    catalog.validate_params("ppt", action, {**params, "overwrite": "yes"})

        self.assertLess(
            script.index("GetActiveObject('Kwpp.Application')"),
            script.index("New-Object -ComObject 'Kwpp.Application'"),
        )
        startup = script[:script.index("function Exec-ping")]
        self.assertNotIn("Presentations.Add", startup)
        self.assertIn("NO_ACTIVE_DOCUMENT", script)
        self.assertNotIn("PptActionsWithoutActivePresentation", script)
        self.assertRegex(
            script,
            r'if \(\$cmd\.requiresActivePresentation -and -not \(Get-PptActivePresentation\)\) \{\s*\$result = @\{\s*success=\$false\s*code="NO_ACTIVE_DOCUMENT"',
        )
        self.assertRegex(
            script,
            re.compile(r"function Exec-createPresentation\(\$p\).*Presentations\.Add", re.DOTALL),
        )
        self.assertRegex(
            script,
            re.compile(r"function Exec-openPresentation\(\$p\).*Presentations\.Open", re.DOTALL),
        )
        self.assertNotRegex(script, r"\.Quit\s*\(")
        self.assertNotIn("Remove-Item", script)
        self.assertIn("function Invoke-PptSafeTargetWrite", script)
        self.assertEqual(4, script.count("Invoke-PptSafeTargetWrite"))
        self.assertIn("TARGET_EXISTS", script)
        self.assertIn("OVERWRITE_NOT_SAFE", script)
        self.assertIn("OVERWRITE_RESTORE_FAILED", script)
        self.assertIn("[System.IO.File]::Copy($backupPath, $fullPath, $true)", script)
        self.assertEqual(120, wps_ppt.EXEC_TIMEOUT)

        close_contract = catalog.get("ppt", "closePresentation")
        self.assertNotIn("name", close_contract["parameters"]["properties"])
        with self.assertRaises(ActionValidationError):
            catalog.validate_params("ppt", "closePresentation", {"name": "other.pptx"})
        self.assertRegex(
            script,
            re.compile(
                r"function Exec-closePresentation\(\$p\)\s*\{\s*\$pres = Get-ActivePres",
                re.DOTALL,
            ),
        )

    def test_ppt_write_results_use_stable_pre_effect_values_and_one_based_ids(self):
        script = wps_ppt.PS_BRIDGE_SCRIPT

        self.assertIn('$closedName = $pres.Name', script)
        self.assertIn('[void]$pres.Close()', script)
        self.assertIn('data=@{closed=$closedName}', script)
        self.assertNotIn('$newSlide[0].SlideIndex', script)
        self.assertIn('$newSlides.Item(1).SlideIndex', script)
        self.assertNotIn('$ns[0].Id', script)
        self.assertIn('$newShapes.Item(1).Id', script)
        self.assertNotIn('effectId=$eff.EntryEffect', script)
        self.assertIn('effectId=$slide.TimeLine.MainSequence.Count', script)

    def test_ppt_targeting_and_shape_operations_use_wps_object_paths(self):
        script = wps_ppt.PS_BRIDGE_SCRIPT

        self.assertNotIn('$pres.Activate()', script)
        self.assertIn('$pres.Windows.Item(1).Activate()', script)
        self.assertIn('$active = Get-ActivePres', script)
        self.assertIn('$global:ppt.ActiveWindow.Selection.TextRange.Text', script)
        self.assertNotIn('$global:ppt.Selection.Text', script)
        self.assertIn('$slide.Shapes.Range().Align($align, 0)', script)
        self.assertIn(
            '$slide.Shapes.Range().Distribute((Convert-Distribution $p.distribute), 0)',
            script,
        )

    def test_ppt_3d_background_and_footer_actions_use_slide_level_objects(self):
        script = wps_ppt.PS_BRIDGE_SCRIPT

        self.assertIn(
            'return [int]($r -bor ($g -shl 8) -bor ($b -shl 16))',
            script,
        )
        self.assertIn('$slide.Shapes.AddTextEffect(', script)
        self.assertNotIn('$tb.TextEffect.PresetThreeDFormat', script)
        self.assertIn('$slide.FollowMasterBackground = $false', script)
        self.assertIn('[void]$fill.TwoColorGradient(1, 1)', script)
        self.assertNotIn('$slide.Background.Fill.Type =', script)
        self.assertIn('$slide.HeadersFooters.Footer.Visible = $true', script)
        self.assertIn('$slide.HeadersFooters.Footer.Text = $p.text', script)
        self.assertIn('$slide.HeadersFooters.DateAndTime.Visible = $true', script)
        self.assertIn('$slide.HeadersFooters.DateAndTime.Text = $p.text', script)
        self.assertIn('$slide.HeadersFooters.SlideNumber.Visible = $show', script)
        self.assertNotIn('$pres.Footers', script)
        self.assertNotIn('$pres.SlideNumber', script)

    def test_word_contracts_preserve_active_document_and_targets(self):
        catalog = ActionCatalog.from_path()
        script = wps_word.PS_BRIDGE_SCRIPT

        for action, path_field in (
            ("saveAs", "filePath"),
            ("convertToPDF", "outputPath"),
            ("convertFormat", "outputPath"),
        ):
            with self.subTest(action=action):
                contract = catalog.get("word", action)
                self.assertEqual("destructive", contract["risk"])
                self.assertEqual(
                    "boolean", contract["parameters"]["properties"]["overwrite"]["type"],
                )
                params = {"targetFormat": "docx"} if action == "convertFormat" else {}
                if action == "saveAs":
                    params[path_field] = r"C:\\tmp\\report.docx"
                catalog.validate_params("word", action, {**params, "overwrite": True})
                with self.assertRaises(ActionValidationError):
                    catalog.validate_params("word", action, {**params, "overwrite": "yes"})

        self.assertLess(
            script.index("GetActiveObject('Kwps.Application')"),
            script.index("New-Object -ComObject 'Kwps.Application'"),
        )
        startup = script[:script.index("function Exec-ping")]
        self.assertNotIn("Documents.Add", startup)
        self.assertIn("NO_ACTIVE_DOCUMENT", script)
        self.assertRegex(
            script,
            r'if \(\$cmd\.requiresActiveDocument -and -not \(Get-WordActiveDocument\)\) \{\s*\$result = @\{\s*success=\$false\s*code="NO_ACTIVE_DOCUMENT"',
        )
        self.assertRegex(
            script,
            re.compile(r"function Exec-createDocument\(\$p\).*Documents\.Add", re.DOTALL),
        )
        self.assertRegex(
            script,
            re.compile(r"function Exec-openDocument\(\$p\).*Documents\.Open", re.DOTALL),
        )
        self.assertNotRegex(script, r"\.Quit\s*\(")
        self.assertNotIn("Remove-Item", script)
        self.assertIn("function Invoke-WordSafeTargetWrite", script)
        self.assertEqual(4, script.count("Invoke-WordSafeTargetWrite"))
        self.assertIn("TARGET_EXISTS", script)
        self.assertIn("OVERWRITE_NOT_SAFE", script)
        self.assertIn("OVERWRITE_RESTORE_FAILED", script)
        self.assertIn("[System.IO.File]::Copy($backupPath, $fullPath, $true)", script)
        self.assertEqual(120, wps_word.EXEC_TIMEOUT)

    def test_word_switch_bookmark_and_style_paths_preserve_document_state(self):
        script = wps_word.PS_BRIDGE_SCRIPT

        self.assertIn(
            'return [int]($r -bor ($g -shl 8) -bor ($b -shl 16))',
            script,
        )
        self.assertNotIn('$doc.Activate()', script)
        self.assertIn('$doc.Windows.Item(1).Activate()', script)
        self.assertIn('$active = Get-ActiveDoc', script)
        self.assertIn('切换文档后活动目标不匹配', script)
        self.assertIn('$rng = $doc.Range($insertAt, $insertAt)', script)
        self.assertIn('$start = [int]$bookmark.Range.Start', script)
        self.assertIn('$end = [int]$bookmark.Range.End', script)
        self.assertIn('$replacement = $doc.Range($start, $end)', script)
        self.assertIn('[void]$doc.Bookmarks.Add($p.name, $bookmarkRange)', script)
        self.assertIn('function Resolve-WordStyle($doc, $name)', script)
        self.assertIn('"heading 1" { return $doc.Styles.Item(-2) }', script)
        self.assertIn('$rng.set_Style($style)', script)


if __name__ == "__main__":
    unittest.main()
