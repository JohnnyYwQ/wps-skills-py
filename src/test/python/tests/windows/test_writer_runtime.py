from pathlib import Path
import unittest
from unittest.mock import patch

from wps_skills.windows.writer_runtime import (
    LazyWindowsWriterBridge,
    windows_powershell_executable,
)


class RecordingLauncher:
    def __init__(self):
        self.commands = []
        self.process = object()

    def start_bridge(self, command):
        self.commands.append(tuple(command))
        return self.process


class RecordingBridge:
    def __init__(self, *, transport):
        self.transport = transport
        self.calls = []

    def execute(self, operation, arguments, context):
        self.calls.append((operation, arguments, context))
        return {"ok": True}

    def close(self):
        return True


class LazyWindowsWriterBridgeTests(unittest.TestCase):
    @patch("wps_skills.windows.writer_runtime.Path.is_file", autospec=True)
    @patch.dict(
        "wps_skills.windows.writer_runtime.os.environ",
        {"WINDIR": r"C:\Windows"},
    )
    def test_native_windows_powershell_is_preferred(self, is_file):
        is_file.return_value = True

        executable = windows_powershell_executable()

        expected = (
            Path(r"C:\Windows")
            / "System32"
            / "WindowsPowerShell"
            / "v1.0"
            / "powershell.exe"
        )
        self.assertEqual(str(expected), executable)
        is_file.assert_called_once_with(expected)

    @patch("wps_skills.windows.writer_runtime.Path.is_file", autospec=True)
    @patch.dict(
        "wps_skills.windows.writer_runtime.os.environ",
        {"WINDIR": r"C:\Windows"},
    )
    def test_missing_native_windows_powershell_fails(self, is_file):
        is_file.return_value = False

        system32 = (
            Path(r"C:\Windows")
            / "System32"
            / "WindowsPowerShell"
            / "v1.0"
            / "powershell.exe"
        )
        with self.assertRaisesRegex(
            FileNotFoundError,
            "Windows PowerShell could not be located",
        ):
            windows_powershell_executable()
        is_file.assert_called_once_with(system32)

    def _execute(self, *, debug_close_created_document):
        launcher = RecordingLauncher()
        bridge = LazyWindowsWriterBridge(
            launcher=launcher,
            script_path=Path("writer_bridge.ps1"),
            executable_factory=lambda: "powershell.exe",
            transport_factory=lambda *, process: ("transport", process),
            bridge_factory=RecordingBridge,
            debug_close_created_document=debug_close_created_document,
        )

        result = bridge.execute("probe", {}, None)

        self.assertEqual({"ok": True}, result)
        return launcher.commands[0]

    def test_normal_session_does_not_request_document_cleanup(self):
        command = self._execute(debug_close_created_document=False)

        window_style = command.index("-WindowStyle")
        self.assertEqual("Hidden", command[window_style + 1])
        self.assertNotIn("-DebugCloseCreatedDocument", command)

    def test_debug_session_explicitly_requests_created_document_cleanup(self):
        command = self._execute(debug_close_created_document=True)

        self.assertEqual("-DebugCloseCreatedDocument", command[-1])

    def test_debug_cleanup_flag_must_be_boolean(self):
        with self.assertRaisesRegex(TypeError, "must be Boolean"):
            LazyWindowsWriterBridge(
                launcher=RecordingLauncher(),
                script_path=Path("writer_bridge.ps1"),
                debug_close_created_document="yes",
            )


if __name__ == "__main__":
    unittest.main()
