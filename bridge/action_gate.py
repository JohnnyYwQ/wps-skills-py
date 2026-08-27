"""Cross-process serialization for one Windows Action at a time."""

from __future__ import annotations

import ctypes
from ctypes import wintypes
import math


ACTION_MUTEX_NAME = r"Local\WpsSkills.ActionRuntime"

_WAIT_OBJECT_0 = 0x00000000
_WAIT_ABANDONED = 0x00000080
_WAIT_TIMEOUT = 0x00000102
_INFINITE = 0xFFFFFFFF


class WindowsActionGate:
    """A named mutex scoped to the current Windows user session.

    This is deliberately only a kernel synchronization primitive.  It does
    not encode a queue or promise any business ordering between Action calls.
    """

    def __init__(self, name=ACTION_MUTEX_NAME):
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.argtypes = (wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)
        kernel32.CreateMutexW.restype = wintypes.HANDLE
        kernel32.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
        kernel32.WaitForSingleObject.restype = wintypes.DWORD
        kernel32.ReleaseMutex.argtypes = (wintypes.HANDLE,)
        kernel32.ReleaseMutex.restype = wintypes.BOOL
        kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        kernel32.CloseHandle.restype = wintypes.BOOL
        self._kernel32 = kernel32
        self._handle = kernel32.CreateMutexW(None, False, name)
        if not self._handle:
            raise ctypes.WinError(ctypes.get_last_error())
        self._acquired = False

    def acquire(self, timeout_seconds):
        milliseconds = _INFINITE if timeout_seconds is None else min(
            _INFINITE - 1,
            max(0, math.ceil(timeout_seconds * 1000)),
        )
        outcome = self._kernel32.WaitForSingleObject(self._handle, milliseconds)
        if outcome in (_WAIT_OBJECT_0, _WAIT_ABANDONED):
            self._acquired = True
            return True
        if outcome == _WAIT_TIMEOUT:
            return False
        raise ctypes.WinError(ctypes.get_last_error())

    def release(self):
        if self._acquired:
            if not self._kernel32.ReleaseMutex(self._handle):
                raise ctypes.WinError(ctypes.get_last_error())
            self._acquired = False

    def close(self):
        if self._handle:
            self._kernel32.CloseHandle(self._handle)
            self._handle = None
