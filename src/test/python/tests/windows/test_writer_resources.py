import re
import unittest

from wps_skills.cli.call import WRITER_BRIDGE_SCRIPT
from wps_skills.word.handlers import _PRIVATE_OPERATIONS


class WriterPowerShellResourceTests(unittest.TestCase):
    def test_bridge_emits_ascii_safe_json_across_windows_code_pages(self):
        bridge_source = WRITER_BRIDGE_SCRIPT.read_text(encoding="utf-8")

        self.assertIn("$json.ToCharArray()", bridge_source)
        self.assertIn("$ascii.AppendFormat('\\u{0:x4}'", bridge_source)
        self.assertIn(
            "[Console]::Out.WriteLine($ascii.ToString())",
            bridge_source,
        )

    def test_every_required_handler_operation_has_a_bridge_dispatch_case(self):
        bridge_source = WRITER_BRIDGE_SCRIPT.read_text(encoding="utf-8")
        dispatch_operations = set(re.findall(
            r"^\s{12}'([^']+)'\s*\{",
            bridge_source,
            flags=re.MULTILINE,
        ))

        self.assertTrue(
            set(_PRIVATE_OPERATIONS.values()).issubset(dispatch_operations)
        )

    def test_bridge_requires_and_loads_the_separate_action_resource(self):
        bridge_source = WRITER_BRIDGE_SCRIPT.read_text(encoding="utf-8")
        actions_path = WRITER_BRIDGE_SCRIPT.with_name("writer_actions.ps1")

        self.assertTrue(actions_path.is_file())
        self.assertIn(
            "$writerActionsPath = Join-Path $PSScriptRoot "
            "'writer_actions.ps1'",
            bridge_source,
        )
        self.assertIn(". $writerActionsPath", bridge_source)

    def test_body_text_normalization_covers_live_wps_table_markers(self):
        bridge_source = WRITER_BRIDGE_SCRIPT.read_text(encoding="utf-8")

        # Live WPS 12 exposes a cell terminator as CR+BEL and adds a second
        # terminator at each row boundary.  Their order matters: normalize the
        # doubled row marker before the single cell marker.
        row_marker = 'Replace("`r`a`r`a", "`n")'
        cell_marker = 'Replace("`r`a", "`t")'
        self.assertLess(
            bridge_source.index(row_marker),
            bridge_source.index(cell_marker),
        )

    def test_image_and_section_break_results_use_observed_wps_facts(self):
        actions_source = WRITER_BRIDGE_SCRIPT.with_name(
            "writer_actions.ps1"
        ).read_text(encoding="utf-8")

        self.assertIn(
            "$observedAlt = [string]$image.AlternativeText",
            actions_source,
        )
        self.assertIn("$wrapFormat = $image.WrapFormat", actions_source)
        self.assertIn("$resultImage['wrap'] = $observedWrap", actions_source)
        self.assertIn(
            "relativeTo = $observedHorizontalReference",
            actions_source,
        )
        self.assertIn(
            "relativeTo = $observedVerticalReference",
            actions_source,
        )
        self.assertIn(
            "$sectionStart = [int]$pageSetup.SectionStart",
            actions_source,
        )
        self.assertIn("type = $observedBreakType", actions_source)

    def test_text_format_maps_script_fonts_and_reads_them_back(self):
        bridge_source = WRITER_BRIDGE_SCRIPT.read_text(encoding="utf-8")

        self.assertIn(
            "$font.NameAscii = [string]$Patch.westernFontFamily",
            bridge_source,
        )
        self.assertIn(
            "$font.NameOther = [string]$Patch.westernFontFamily",
            bridge_source,
        )
        self.assertIn(
            "$font.NameFarEast = [string]$Patch.eastAsiaFontFamily",
            bridge_source,
        )
        self.assertIn(
            "$westernFontName = [string]$font.NameAscii",
            bridge_source,
        )
        self.assertIn(
            "$eastAsiaFontName = [string]$font.NameFarEast",
            bridge_source,
        )


if __name__ == "__main__":
    unittest.main()
