import sys
import unittest
from unittest.mock import patch

from windows_com import (
    ComRegistration,
    NativeWindowsComEnvironment,
    resolve_com_runtime,
)


class _FakeWindowsEnvironment:
    def __init__(self, registrations, process_bitness=64, os_bitness=64):
        self._registrations = registrations
        self.process_bitness = process_bitness
        self.os_bitness = os_bitness

    def inspect_registration(self, progid, view_bitness):
        return self._registrations[view_bitness]

    def powershell_executable(self, view_bitness):
        return {
            32: r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe",
            64: r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        }[view_bitness]


class _PartiallyUnreadableEnvironment(_FakeWindowsEnvironment):
    def inspect_registration(self, progid, view_bitness):
        if view_bitness == 64:
            raise PermissionError("registry access denied")
        return super().inspect_registration(progid, view_bitness)


class _FakeRegistryKey:
    def __init__(self, registry, path, view_bitness):
        self.registry = registry
        self.path = path
        self.view_bitness = view_bitness

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return False


class _FakeWinreg:
    HKEY_CLASSES_ROOT = object()
    KEY_READ = 0x0001
    KEY_WOW64_32KEY = 0x0200
    KEY_WOW64_64KEY = 0x0100

    def __init__(self, values):
        self._values = values

    def OpenKey(self, root, path, reserved, access):
        view_bitness = 32 if access & self.KEY_WOW64_32KEY else 64
        if not any(
            view == view_bitness and registered_path == path
            for view, registered_path, _ in self._values
        ):
            raise FileNotFoundError(path)
        return _FakeRegistryKey(self, path, view_bitness)

    def QueryValue(self, key, sub_key):
        return self._query(key, "")

    def QueryValueEx(self, key, value_name):
        return self._query(key, value_name), 1

    def _query(self, key, value_name):
        try:
            return self._values[(key.view_bitness, key.path, value_name)]
        except KeyError as exc:
            raise FileNotFoundError(value_name) from exc


class WindowsComResolutionTests(unittest.TestCase):
    def test_only_complete_32_bit_registration_selects_32_bit_powershell(self):
        progid = "Ket.Application"
        environment = _FakeWindowsEnvironment({
            32: ComRegistration(
                progid=progid,
                view_bitness=32,
                clsid="{45540001-5750-5300-4B49-4E47534F4655}",
                activation_kind="LocalServer32",
                activation_value=r'C:\Users\tester\wps.exe /Automation',
            ),
            64: ComRegistration(
                progid=progid,
                view_bitness=64,
                clsid="{45540001-5750-5300-4B49-4E47534F4655}",
                problem="CLSID has no activation server",
            ),
        })

        resolution = resolve_com_runtime(progid, environment=environment)

        self.assertTrue(resolution.available)
        self.assertEqual(32, resolution.selected_view_bitness)
        self.assertEqual(
            r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe",
            resolution.powershell_executable,
        )

    def test_both_complete_registrations_prefer_the_current_process_bitness(self):
        progid = "Kwpp.Application"
        environment = _FakeWindowsEnvironment({
            32: ComRegistration(
                progid=progid,
                view_bitness=32,
                clsid="{PPT-32}",
                activation_kind="LocalServer32",
                activation_value=r"C:\WPS32\wpp.exe",
            ),
            64: ComRegistration(
                progid=progid,
                view_bitness=64,
                clsid="{PPT-64}",
                activation_kind="LocalServer32",
                activation_value=r"C:\WPS64\wpp.exe",
            ),
        })

        resolution = resolve_com_runtime(progid, environment=environment)

        self.assertEqual(64, resolution.selected_view_bitness)

    def test_no_complete_registration_returns_actionable_diagnostics(self):
        progid = "Kwps.Application"
        environment = _FakeWindowsEnvironment({
            32: ComRegistration(
                progid=progid,
                view_bitness=32,
                problem="ProgID not found",
            ),
            64: ComRegistration(
                progid=progid,
                view_bitness=64,
                clsid="{WORD}",
                problem="CLSID has no activation server",
            ),
        })

        resolution = resolve_com_runtime(progid, environment=environment)

        self.assertFalse(resolution.available)
        self.assertIn("32-bit: ProgID not found", resolution.diagnostic)
        self.assertIn(
            "64-bit: CLSID {WORD}; CLSID has no activation server",
            resolution.diagnostic,
        )

    def test_32_bit_python_uses_sysnative_for_a_64_bit_registry_view(self):
        environment = NativeWindowsComEnvironment(
            process_bitness=32,
            os_bitness=64,
            windows_directory=r"C:\Windows",
        )

        executable = environment.powershell_executable(64)

        self.assertEqual(
            r"C:\Windows\Sysnative\WindowsPowerShell\v1.0\powershell.exe",
            executable,
        )

    def test_one_unreadable_view_does_not_hide_an_activatable_other_view(self):
        progid = "Ket.Application"
        environment = _PartiallyUnreadableEnvironment({
            32: ComRegistration(
                progid=progid,
                view_bitness=32,
                clsid="{EXCEL}",
                activation_kind="LocalServer32",
                activation_value=r"C:\WPS\wps.exe",
            ),
        })

        resolution = resolve_com_runtime(progid, environment=environment)

        self.assertTrue(resolution.available)
        self.assertEqual(32, resolution.selected_view_bitness)
        self.assertIn("registry access denied", resolution.diagnostic)

    def test_native_registry_adapter_reads_each_wow64_view_independently(self):
        progid = "Ket.Application"
        clsid = "{45540001-5750-5300-4B49-4E47534F4655}"
        fake_winreg = _FakeWinreg({
            (32, r"Ket.Application\CLSID", ""): clsid,
            (32, rf"CLSID\{clsid}\LocalServer32", ""): r"C:\WPS\wps.exe",
            (64, r"Ket.Application\CLSID", ""): clsid,
        })
        environment = NativeWindowsComEnvironment(
            process_bitness=64,
            os_bitness=64,
            windows_directory=r"C:\Windows",
        )

        with patch.dict(sys.modules, {"winreg": fake_winreg}):
            registration_32 = environment.inspect_registration(progid, 32)
            registration_64 = environment.inspect_registration(progid, 64)

        self.assertTrue(registration_32.activatable)
        self.assertEqual("LocalServer32", registration_32.activation_kind)
        self.assertFalse(registration_64.activatable)
        self.assertEqual(
            "CLSID has no activation server",
            registration_64.problem,
        )


if __name__ == "__main__":
    unittest.main()
