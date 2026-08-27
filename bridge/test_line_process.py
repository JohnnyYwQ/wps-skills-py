import subprocess
import unittest
from unittest.mock import Mock

from line_process import stop_line_process


class LineProcessShutdownTests(unittest.TestCase):
    def test_timeout_escalates_from_exit_to_terminate_then_kill(self):
        process = Mock()
        process.poll.return_value = None
        process.wait.side_effect = [
            subprocess.TimeoutExpired("powershell", 2),
            subprocess.TimeoutExpired("powershell", 1),
            0,
        ]

        stop_line_process(process, graceful_timeout=2, forced_timeout=1)

        process.stdin.write.assert_called_once_with("EXIT\n")
        process.stdin.flush.assert_called_once_with()
        process.terminate.assert_called_once_with()
        process.kill.assert_called_once_with()
        self.assertEqual(3, process.wait.call_count)


if __name__ == "__main__":
    unittest.main()
