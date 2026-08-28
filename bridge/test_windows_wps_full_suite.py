"""Cross-platform checks for the Windows real-WPS full Action suite."""

import sys
import subprocess
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))
if str(ROOT / "bridge") not in sys.path:
    sys.path.insert(0, str(ROOT / "bridge"))

from action_catalog import ActionCatalog, load_action_manifest  # noqa: E402
from windows_wps_full_suite import (  # noqa: E402
    build_action_params,
    default_suite_state,
    iter_declared_actions,
)
from test_windows_wps_all_actions import _effect_errors  # noqa: E402


class WindowsWpsFullSuiteTest(unittest.TestCase):
    def test_every_action_contract_is_declared_exactly_once(self):
        manifest_keys = {
            (contract["owner"], contract["action"])
            for contract in load_action_manifest()["actions"]
        }
        declared = list(iter_declared_actions())

        self.assertEqual(235, len(manifest_keys))
        self.assertEqual(len(declared), len(set(declared)), "suite has duplicate Actions")
        self.assertEqual(manifest_keys, set(declared))

    def test_every_declared_case_has_contract_valid_default_params(self):
        catalog = ActionCatalog.from_path()
        state = default_suite_state(ROOT / "test-results" / "dry-run")

        for owner, action in iter_declared_actions():
            with self.subTest(owner=owner, action=action):
                params = build_action_params(owner, action, state)
                catalog.validate_params(owner, action, params)

    def test_runner_validate_only_mode_is_cross_platform(self):
        completed = subprocess.run(
            [
                sys.executable,
                str(SCRIPTS / "test_windows_wps_all_actions.py"),
                "--validate-only",
            ],
            cwd=ROOT,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            check=False,
        )

        self.assertEqual(0, completed.returncode, completed.stdout)
        self.assertIn("235 Action cases", completed.stdout)

    def test_windows_entrypoint_runs_parse_check_then_full_runner(self):
        entrypoint = SCRIPTS / "test_windows_wps_all_actions.ps1"
        script = entrypoint.read_text(encoding="utf-8")

        self.assertIn("test_windows_powershell_parse.ps1", script)
        self.assertIn("test_windows_wps_all_actions.py", script)
        self.assertIn("-X", script)
        self.assertIn("utf8", script)

    def test_comment_effect_validation_uses_the_contract_comments_array(self):
        errors = _effect_errors(
            "excel",
            "getCellComments",
            {"range": "R1:R2"},
            {"success": True, "data": {"comments": [{"cell": "$R$1", "text": "ok"}]}},
        )

        self.assertEqual([], errors)

    def test_slide_title_readback_runs_after_the_title_write(self):
        declared = list(iter_declared_actions())

        self.assertLess(
            declared.index(("ppt", "setSlideTitle")),
            declared.index(("ppt", "getSlideTitle")),
        )

    def test_stateful_write_readbacks_run_after_their_writes(self):
        declared = list(iter_declared_actions())

        for owner, write, readback in (
            ("excel", "copySheet", "getSheetList"),
            ("ppt", "addAnimation", "getAnimations"),
            ("ppt", "addAnimationPreset", "getAnimations"),
            ("ppt", "addEmphasisAnimation", "getAnimations"),
            ("word", "switchDocument", "getActiveDocument"),
            ("word", "replaceBookmarkContent", "getDocumentText"),
        ):
            with self.subTest(owner=owner, write=write, readback=readback):
                self.assertLess(
                    declared.index((owner, write)),
                    declared.index((owner, readback)),
                )

    def test_stateful_effect_validation_rejects_missing_write_effects(self):
        missing_copy = _effect_errors(
            "excel",
            "getSheetList",
            {},
            {"success": True, "data": {"sheets": [{"name": "SuiteRenamed"}]}},
        )
        intact_word = _effect_errors(
            "word",
            "getDocumentText",
            {},
            {"success": True, "data": {"text": "WORD_FIND_TOKEN WORD_BOOKMARK_TOKEN"}},
        )
        missing_bookmark = _effect_errors(
            "word",
            "getDocumentText",
            {},
            {"success": True, "data": {"text": "WORD_FIND_TOKEN"}},
        )
        wrong_active_doc = _effect_errors(
            "word",
            "getActiveDocument",
            {},
            {"success": True, "data": {"name": "word-secondary.docx"}},
        )
        too_few_animations = _effect_errors(
            "ppt",
            "getAnimations",
            {},
            {"success": True, "data": {"count": 2, "animations": [{}, {}]}},
        )

        self.assertTrue(missing_copy)
        self.assertEqual([], intact_word)
        self.assertTrue(missing_bookmark)
        self.assertTrue(wrong_active_doc)
        self.assertTrue(too_few_animations)


if __name__ == "__main__":
    unittest.main()
