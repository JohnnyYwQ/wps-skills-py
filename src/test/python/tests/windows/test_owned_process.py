import io
import subprocess
import unittest

from wps_skills.windows.owned_process import (
    ProcessOwnershipUnavailable,
    WindowsOwnedProcessLauncher,
)


class FakeJobApi:
    def __init__(self, *, assign_error=None):
        self.assign_error = assign_error
        self.calls = []

    def create_job(self):
        self.calls.append(("create",))
        return "job-1"

    def assign(self, job, process):
        self.calls.append(("assign", job, process.pid))
        if self.assign_error is not None:
            raise self.assign_error

    def resume(self, process):
        self.calls.append(("resume", process.pid))

    def close_job(self, job):
        self.calls.append(("close", job))


class ClosingInput(io.StringIO):
    def __init__(self, process, *, exits):
        super().__init__()
        self._process = process
        self._exits = exits

    def close(self):
        super().close()
        if self._exits:
            self._process.returncode = 0


class FakeProcess:
    def __init__(self, pid, *, graceful=True):
        self.pid = pid
        self._handle = pid + 1000
        self.returncode = None
        self.stdin = ClosingInput(self, exits=graceful)
        self.stdout = io.StringIO()
        self.stderr = io.StringIO()
        self.kill_calls = 0
        self.terminate_calls = 0

    def poll(self):
        return self.returncode

    def wait(self, timeout):
        if self.returncode is None:
            raise TimeoutError
        return self.returncode

    def terminate(self):
        self.terminate_calls += 1
        self.returncode = 1

    def kill(self):
        self.kill_calls += 1
        self.returncode = 1


class PopenFactory:
    def __init__(self, process):
        self.process = process
        self.calls = []

    def __call__(self, command, **kwargs):
        self.calls.append((command, kwargs))
        return self.process


class WindowsOwnedProcessLauncherTests(unittest.TestCase):
    def test_assigns_suspended_process_before_resume_and_reports_cleanup(self):
        job = FakeJobApi()
        process = FakeProcess(42)
        popen = PopenFactory(process)
        launcher = WindowsOwnedProcessLauncher(
            job_api=job,
            popen_factory=popen,
        )

        returned = launcher.start_bridge(["powershell.exe", "-File", "bridge.ps1"])
        cleanup = launcher.close()

        self.assertIs(process, returned)
        self.assertEqual("assign", job.calls[1][0])
        self.assertEqual("resume", job.calls[2][0])
        creation_flags = popen.calls[0][1]["creationflags"]
        self.assertTrue(
            creation_flags
            & getattr(subprocess, "CREATE_SUSPENDED", 0x00000004)
        )
        self.assertTrue(
            creation_flags
            & getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        )
        self.assertEqual(42, cleanup[0].pid)
        self.assertEqual(("graceful",), cleanup[0].cleanup_steps)
        self.assertTrue(cleanup[0].released)

    def test_assignment_failure_kills_the_still_suspended_child(self):
        job = FakeJobApi(assign_error=OSError("cannot assign"))
        process = FakeProcess(43, graceful=False)
        launcher = WindowsOwnedProcessLauncher(
            job_api=job,
            popen_factory=PopenFactory(process),
        )

        with self.assertRaises(ProcessOwnershipUnavailable):
            launcher.start_bridge(["powershell.exe", "-File", "bridge.ps1"])

        self.assertEqual(1, process.kill_calls)
        self.assertFalse(any(call[0] == "resume" for call in job.calls))
        self.assertEqual((), launcher.close())


if __name__ == "__main__":
    unittest.main()
