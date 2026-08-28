import subprocess
import sys
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class ZeroInstallDeliveryTests(unittest.TestCase):
    def test_environment_check_reports_without_package_installation_guidance(self):
        result = subprocess.run(
            [sys.executable, "scripts/install.py", "--check"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
        )

        self.assertEqual(0, result.returncode, result.stderr)
        output = result.stdout.lower()
        self.assertNotIn("mcp", output)
        self.assertNotIn("pip", output)
        self.assertNotIn("venv", output)

    def test_distribution_has_no_legacy_protocol_or_service_entrypoints(self):
        for relative_path in (
            "bridge/mcp_adapter.py",
            "bridge/server.py",
            "bridge/service_lifecycle.py",
            "scripts/mcp_server.py",
            "scripts/service.py",
            "scripts/start.py",
            "scripts/test.py",
            "requirements.txt",
        ):
            self.assertFalse((ROOT / relative_path).exists(), relative_path)

    def test_legacy_service_troubleshooting_documents_are_not_distributed(self):
        self.assertFalse((ROOT / "docs/bridge-health-timeout-handoff.md").exists())

    def test_user_facing_documents_only_describe_skill_scoped_cli_workflow(self):
        prohibited_terms = ("mcp", "http", "health", "service.py", "restart", "pip", "venv")
        for relative_path in ("SKILL.md", "README.md", "INSTALL.md"):
            content = (ROOT / relative_path).read_text(encoding="utf-8").lower()
            for term in prohibited_terms:
                self.assertNotIn(term, content, f"{relative_path} still mentions {term}")

    def test_windows_powershell_entry_scripts_are_ascii_for_powershell_5_1(self):
        for relative_path in (
            "scripts/test_windows_powershell_parse.ps1",
            "scripts/test_windows_wps_regressions.ps1",
        ):
            with self.subTest(path=relative_path):
                content = (ROOT / relative_path).read_bytes()
                try:
                    content.decode("ascii")
                except UnicodeDecodeError as exc:
                    self.fail(
                        f"{relative_path} must be ASCII because Windows PowerShell 5.1 "
                        f"treats UTF-8 without BOM as the system ANSI code page: {exc}"
                    )


if __name__ == "__main__":
    unittest.main()
