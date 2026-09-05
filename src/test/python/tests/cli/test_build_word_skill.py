import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from wps_skills.cli.build_word_skill import build_word_skill
from wps_skills.word.contracts import WORD_PRODUCTION_CONTRACT_SET


class WordSkillBuildTests(unittest.TestCase):
    def test_relocated_skill_discovers_and_resolves_without_repository_or_windows(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            built = build_word_skill(root / "first" / "wps-word")
            installed = root / "installed" / "wps-word"
            installed.parent.mkdir()
            built.rename(installed)
            environment = dict(os.environ, PYTHONPATH="", PYTHONNOUSERSITE="1")
            entry = installed / "scripts" / "word.py"
            result = subprocess.run([sys.executable, str(entry), "--app", "word", "--index"], cwd=root, env=environment, capture_output=True, text=True, check=True)
            index = json.loads(result.stdout)
            self.assertEqual([entry.to_wire() for entry in WORD_PRODUCTION_CONTRACT_SET.action_index()], index["actions"])
            result = subprocess.run([sys.executable, str(entry), "--app", "word", "--resolve", "createDocument", "writeContent", "inspectDocument"], cwd=root, env=environment, capture_output=True, text=True, check=True)
            self.assertEqual("complete", json.loads(result.stdout)["status"])
            # The packaged Runtime must also resolve the actual PowerShell resources.
            code = "import sys; sys.path.insert(0, sys.argv[1]); import word; from wps_skills.cli.call import WRITER_BRIDGE_SCRIPT; assert WRITER_BRIDGE_SCRIPT.is_file(); assert WRITER_BRIDGE_SCRIPT.with_name('writer_actions.ps1').is_file(); c=word.open_session(); assert not c.can_execute"
            subprocess.run([sys.executable, "-c", code, str(entry.parent)], cwd=root, env=environment, check=True)
            manifest = json.loads((installed / "runtime/files.sha256.json").read_text())
            for name, digest in manifest.items():
                self.assertEqual(digest, hashlib.sha256((installed / name).read_bytes()).hexdigest())
            self.assertTrue((installed / "SKILL.md").is_file())
            self.assertFalse((installed / "runtime/src/test").exists())
            self.assertFalse((installed / "runtime/src/main/resources/skills").exists())

    def test_existing_destination_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as temporary:
            destination = Path(temporary) / "wps-word"
            destination.mkdir()
            marker = destination / "user.txt"
            marker.write_text("keep")
            with self.assertRaises(FileExistsError):
                build_word_skill(destination)
            self.assertEqual("keep", marker.read_text())


if __name__ == "__main__":
    unittest.main()
