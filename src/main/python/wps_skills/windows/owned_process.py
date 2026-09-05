"""Session-owned Windows process launcher backed by one Job Object."""

from collections import deque
import ctypes
from ctypes import wintypes
import os
import subprocess
import threading
import time

from wps_skills.core.action_session import ProcessCleanup


class ProcessOwnershipUnavailable(RuntimeError):
    """A child could not be placed under the Session Job Object."""


class _IoCounters(ctypes.Structure):
    _fields_ = [
        ("ReadOperationCount", ctypes.c_ulonglong),
        ("WriteOperationCount", ctypes.c_ulonglong),
        ("OtherOperationCount", ctypes.c_ulonglong),
        ("ReadTransferCount", ctypes.c_ulonglong),
        ("WriteTransferCount", ctypes.c_ulonglong),
        ("OtherTransferCount", ctypes.c_ulonglong),
    ]


class _BasicLimitInformation(ctypes.Structure):
    _fields_ = [
        ("PerProcessUserTimeLimit", ctypes.c_longlong),
        ("PerJobUserTimeLimit", ctypes.c_longlong),
        ("LimitFlags", wintypes.DWORD),
        ("MinimumWorkingSetSize", ctypes.c_size_t),
        ("MaximumWorkingSetSize", ctypes.c_size_t),
        ("ActiveProcessLimit", wintypes.DWORD),
        ("Affinity", ctypes.c_size_t),
        ("PriorityClass", wintypes.DWORD),
        ("SchedulingClass", wintypes.DWORD),
    ]


class _ExtendedLimitInformation(ctypes.Structure):
    _fields_ = [
        ("BasicLimitInformation", _BasicLimitInformation),
        ("IoInfo", _IoCounters),
        ("ProcessMemoryLimit", ctypes.c_size_t),
        ("JobMemoryLimit", ctypes.c_size_t),
        ("PeakProcessMemoryUsed", ctypes.c_size_t),
        ("PeakJobMemoryUsed", ctypes.c_size_t),
    ]


class _WindowsJobApi:
    _EXTENDED_LIMIT_INFORMATION = 9
    _KILL_ON_JOB_CLOSE = 0x00002000
    _SILENT_BREAKAWAY_OK = 0x00001000

    def __init__(self):
        if os.name != "nt":
            raise OSError("Windows Job Objects are available only on Windows")
        self._kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        self._ntdll = ctypes.WinDLL("ntdll", use_last_error=True)
        self._kernel32.CreateJobObjectW.restype = wintypes.HANDLE
        self._kernel32.CreateJobObjectW.argtypes = [
            ctypes.c_void_p,
            wintypes.LPCWSTR,
        ]
        self._kernel32.SetInformationJobObject.restype = wintypes.BOOL
        self._kernel32.SetInformationJobObject.argtypes = [
            wintypes.HANDLE,
            ctypes.c_int,
            ctypes.c_void_p,
            wintypes.DWORD,
        ]
        self._kernel32.AssignProcessToJobObject.restype = wintypes.BOOL
        self._kernel32.AssignProcessToJobObject.argtypes = [
            wintypes.HANDLE,
            wintypes.HANDLE,
        ]
        self._kernel32.CloseHandle.restype = wintypes.BOOL
        self._kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
        self._ntdll.NtResumeProcess.restype = ctypes.c_long
        self._ntdll.NtResumeProcess.argtypes = [wintypes.HANDLE]

    @staticmethod
    def _win_error(message):
        return OSError(ctypes.get_last_error(), message)

    def create_job(self):
        handle = self._kernel32.CreateJobObjectW(None, None)
        if not handle:
            raise self._win_error("CreateJobObjectW failed")
        limits = _ExtendedLimitInformation()
        limits.BasicLimitInformation.LimitFlags = (
            self._KILL_ON_JOB_CLOSE | self._SILENT_BREAKAWAY_OK
        )
        configured = self._kernel32.SetInformationJobObject(
            handle,
            self._EXTENDED_LIMIT_INFORMATION,
            ctypes.byref(limits),
            ctypes.sizeof(limits),
        )
        if not configured:
            self._kernel32.CloseHandle(handle)
            raise self._win_error("SetInformationJobObject failed")
        return handle

    def assign(self, job, process):
        if not self._kernel32.AssignProcessToJobObject(
            job,
            wintypes.HANDLE(process._handle),
        ):
            raise self._win_error("AssignProcessToJobObject failed")

    def resume(self, process):
        status = self._ntdll.NtResumeProcess(
            wintypes.HANDLE(process._handle)
        )
        if status != 0:
            raise OSError(status, "NtResumeProcess failed")

    def close_job(self, job):
        if job and not self._kernel32.CloseHandle(job):
            raise self._win_error("CloseHandle(Job) failed")


class WindowsOwnedProcessLauncher:
    """Register every bridge before it runs and reclaim all of them boundedly."""

    def __init__(
        self,
        *,
        job_api=None,
        popen_factory=subprocess.Popen,
        clock=time.monotonic,
        cleanup_timeout_seconds=5,
        graceful_timeout_seconds=2,
        terminate_timeout_seconds=1,
    ):
        if min(
            cleanup_timeout_seconds,
            graceful_timeout_seconds,
            terminate_timeout_seconds,
        ) <= 0:
            raise ValueError("process cleanup timeouts must be positive")
        self._job_api = job_api or _WindowsJobApi()
        self._popen_factory = popen_factory
        self._clock = clock
        self._cleanup_timeout_seconds = cleanup_timeout_seconds
        self._graceful_timeout_seconds = graceful_timeout_seconds
        self._terminate_timeout_seconds = terminate_timeout_seconds
        self._job = self._job_api.create_job()
        self._processes = []
        self._stderr_tails = {}
        self._lock = threading.Lock()
        self._close_lock = threading.Lock()
        self._closed = False
        self._cleanup = None

    @staticmethod
    def _drain_stderr(stream, tail):
        try:
            for line in stream:
                tail.append(line.rstrip("\r\n"))
        except Exception:
            pass

    def start_bridge(self, command):
        if (
            not isinstance(command, (list, tuple))
            or not command
            or any(not isinstance(item, str) or not item for item in command)
        ):
            raise ValueError("bridge command must contain non-empty strings")
        with self._lock:
            if self._closed:
                raise ProcessOwnershipUnavailable("process launcher is closed")
            process = None
            try:
                process = self._popen_factory(
                    list(command),
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding="utf-8",
                    errors="strict",
                    bufsize=1,
                    creationflags=(
                        getattr(subprocess, "CREATE_SUSPENDED", 0x00000004)
                        | getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
                    ),
                )
                self._job_api.assign(self._job, process)
                self._processes.append(process)
                tail = deque(maxlen=64)
                self._stderr_tails[process.pid] = tail
                threading.Thread(
                    target=self._drain_stderr,
                    args=(process.stderr, tail),
                    daemon=True,
                ).start()
                self._job_api.resume(process)
                return process
            except Exception as exc:
                if process is not None:
                    try:
                        process.kill()
                    except Exception:
                        pass
                    try:
                        process.wait(timeout=1)
                    except Exception:
                        pass
                raise ProcessOwnershipUnavailable(
                    "the bridge could not be owned before execution"
                ) from exc

    @staticmethod
    def _wait_until(processes, deadline, clock):
        processes = [
            process for process in processes if process.poll() is None
        ]
        while processes and clock() < deadline:
            processes = [
                process for process in processes if process.poll() is None
            ]
            if processes:
                time.sleep(min(0.02, max(0, deadline - clock())))
        return processes

    def close(self):
        with self._close_lock:
            if self._cleanup is not None:
                return self._cleanup
            with self._lock:
                self._closed = True
                processes = tuple(self._processes)
            steps = {process.pid: [] for process in processes}
            deadline = self._clock() + self._cleanup_timeout_seconds
            running = []
            for process in processes:
                if process.poll() is not None:
                    steps[process.pid].append("already_exited")
                    continue
                steps[process.pid].append("graceful")
                try:
                    process.stdin.close()
                except Exception:
                    pass
                running.append(process)
            graceful_deadline = min(
                deadline,
                self._clock() + self._graceful_timeout_seconds,
            )
            running = self._wait_until(running, graceful_deadline, self._clock)
            for process in running:
                steps[process.pid].append("terminate")
                try:
                    process.terminate()
                except Exception:
                    pass
            terminate_deadline = min(
                deadline,
                self._clock() + self._terminate_timeout_seconds,
            )
            running = self._wait_until(running, terminate_deadline, self._clock)
            for process in running:
                steps[process.pid].append("kill")
                try:
                    process.kill()
                except Exception:
                    pass
            running = self._wait_until(running, deadline, self._clock)
            if running:
                for process in running:
                    steps[process.pid].append("job_close")
            job_closed = True
            try:
                self._job_api.close_job(self._job)
            except Exception:
                job_closed = False
            self._job = None
            running = self._wait_until(running, deadline, self._clock)
            still_running = {process.pid for process in running}
            self._cleanup = tuple(
                ProcessCleanup(
                    pid=process.pid,
                    cleanup_steps=steps[process.pid],
                    released=(process.pid not in still_running and job_closed),
                )
                for process in processes
            )
            if not job_closed and not processes:
                raise ProcessOwnershipUnavailable(
                    "the Session Job Object could not be closed"
                )
            return self._cleanup
