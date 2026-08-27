import io
from contextlib import redirect_stdout
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from scripts import install


class InstallCheckTests(unittest.TestCase):
    def test_windows_check_reports_the_selected_activatable_registry_view(self):
        selected_registration = SimpleNamespace(
            clsid="{45540001-5750-5300-4B49-4E47534F4655}",
            activation_kind="LocalServer32",
        )
        resolution = SimpleNamespace(
            available=True,
            selected_view_bitness=32,
            selected_registration=selected_registration,
            powershell_executable=(
                r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
            ),
            diagnostic="",
        )
        output = io.StringIO()

        with patch.object(install.platform, "system", return_value="Windows"), patch.object(
            install,
            "resolve_com_runtime",
            return_value=resolution,
        ), redirect_stdout(output):
            result = install.check_wps()

        self.assertTrue(result)
        self.assertIn("32-bit view", output.getvalue())
        self.assertIn(selected_registration.clsid, output.getvalue())
        self.assertIn(resolution.powershell_executable, output.getvalue())

    def test_windows_check_rejects_a_progid_without_an_activation_server(self):
        resolution = SimpleNamespace(
            available=False,
            selected_view_bitness=None,
            selected_registration=None,
            powershell_executable=None,
            diagnostic=(
                "32-bit: CLSID {EXCEL}; CLSID has no activation server | "
                "64-bit: ProgID not found"
            ),
        )
        output = io.StringIO()

        with patch.object(install.platform, "system", return_value="Windows"), patch.object(
            install,
            "resolve_com_runtime",
            return_value=resolution,
        ), redirect_stdout(output):
            result = install.check_wps()

        self.assertFalse(result)
        self.assertIn("COM 注册不可激活", output.getvalue())
        self.assertIn("CLSID has no activation server", output.getvalue())

    def test_file_check_no_longer_requires_unused_config_json(self):
        output = io.StringIO()

        with redirect_stdout(output):
            result = install.check_files()

        self.assertTrue(result)
        self.assertNotIn("config.json", output.getvalue())


if __name__ == "__main__":
    unittest.main()
