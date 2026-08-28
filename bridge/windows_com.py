"""Resolve the Windows COM registry view and PowerShell host for WPS."""

from __future__ import annotations

from dataclasses import dataclass
import ntpath
import os
import queue
import struct
from typing import Optional, Protocol, Tuple


@dataclass(frozen=True)
class ComRegistration:
    """COM activation information observed in one registry view."""

    progid: str
    view_bitness: int
    clsid: Optional[str] = None
    activation_kind: Optional[str] = None
    activation_value: Optional[str] = None
    problem: Optional[str] = None

    @property
    def activatable(self) -> bool:
        return bool(self.clsid and self.activation_kind and self.activation_value)


@dataclass(frozen=True)
class ComRuntimeResolution:
    """Selected COM registry view and the PowerShell executable that sees it."""

    progid: str
    process_bitness: int
    os_bitness: int
    registrations: Tuple[ComRegistration, ...]
    selected_view_bitness: Optional[int] = None
    powershell_executable: Optional[str] = None

    @property
    def available(self) -> bool:
        return bool(self.selected_view_bitness and self.powershell_executable)

    @property
    def selected_registration(self) -> Optional[ComRegistration]:
        return next(
            (
                registration
                for registration in self.registrations
                if registration.view_bitness == self.selected_view_bitness
            ),
            None,
        )

    @property
    def diagnostic(self) -> str:
        details = []
        for registration in self.registrations:
            fields = []
            if registration.clsid:
                fields.append(f"CLSID {registration.clsid}")
            if registration.activation_kind:
                fields.append(registration.activation_kind)
            if registration.problem:
                fields.append(registration.problem)
            details.append(
                f"{registration.view_bitness}-bit: "
                + ("; ".join(fields) if fields else "registration unavailable")
            )
        return " | ".join(details)


class WindowsComEnvironment(Protocol):
    process_bitness: int
    os_bitness: int

    def inspect_registration(self, progid: str, view_bitness: int) -> ComRegistration:
        ...

    def powershell_executable(self, view_bitness: int) -> str:
        ...


class NativeWindowsComEnvironment:
    """Adapter over the native Windows registry and WOW64 filesystem layout."""

    def __init__(
        self,
        *,
        process_bitness: Optional[int] = None,
        os_bitness: Optional[int] = None,
        windows_directory: Optional[str] = None,
    ) -> None:
        self.process_bitness = process_bitness or struct.calcsize("P") * 8
        detected_os_bitness = (
            64
            if self.process_bitness == 64 or os.environ.get("PROCESSOR_ARCHITEW6432")
            else 32
        )
        self.os_bitness = os_bitness or detected_os_bitness
        self.windows_directory = (
            windows_directory
            or os.environ.get("WINDIR")
            or os.environ.get("SystemRoot")
        )

    def inspect_registration(self, progid: str, view_bitness: int) -> ComRegistration:
        try:
            import winreg
        except ImportError as exc:
            raise RuntimeError("Windows registry is unavailable on this platform") from exc

        view_flag = (
            winreg.KEY_WOW64_32KEY
            if view_bitness == 32
            else winreg.KEY_WOW64_64KEY
        )
        access = winreg.KEY_READ | view_flag

        def read_default(path: str) -> Optional[str]:
            try:
                with winreg.OpenKey(
                    winreg.HKEY_CLASSES_ROOT,
                    path,
                    0,
                    access,
                ) as key:
                    value = winreg.QueryValue(key, "")
                    return str(value).strip() if value else None
            except FileNotFoundError:
                return None

        clsid = read_default(ntpath.join(progid, "CLSID"))
        if not clsid:
            return ComRegistration(
                progid=progid,
                view_bitness=view_bitness,
                problem="ProgID not found or has no CLSID",
            )

        clsid_path = ntpath.join("CLSID", clsid)
        local_service = None
        try:
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT,
                clsid_path,
                0,
                access,
            ) as clsid_key:
                appid, _ = winreg.QueryValueEx(clsid_key, "AppID")
            if appid:
                try:
                    with winreg.OpenKey(
                        winreg.HKEY_CLASSES_ROOT,
                        ntpath.join("AppID", str(appid)),
                        0,
                        access,
                    ) as appid_key:
                        local_service, _ = winreg.QueryValueEx(
                            appid_key,
                            "LocalService",
                        )
                except (FileNotFoundError, OSError):
                    local_service = None
        except (FileNotFoundError, OSError):
            local_service = None

        if local_service:
            return ComRegistration(
                progid=progid,
                view_bitness=view_bitness,
                clsid=clsid,
                activation_kind="LocalService",
                activation_value=str(local_service),
            )

        for activation_kind in ("LocalServer32", "InprocServer32"):
            activation_value = read_default(ntpath.join(clsid_path, activation_kind))
            if activation_value:
                return ComRegistration(
                    progid=progid,
                    view_bitness=view_bitness,
                    clsid=clsid,
                    activation_kind=activation_kind,
                    activation_value=activation_value,
                )

        return ComRegistration(
            progid=progid,
            view_bitness=view_bitness,
            clsid=clsid,
            problem="CLSID has no activation server",
        )

    def powershell_executable(self, view_bitness: int) -> str:
        if not self.windows_directory:
            raise RuntimeError("WINDIR/SystemRoot is not set")
        if self.os_bitness == 32:
            system_directory = "System32"
        elif view_bitness == 32:
            system_directory = "SysWOW64"
        elif self.process_bitness == 32:
            system_directory = "Sysnative"
        else:
            system_directory = "System32"
        return ntpath.join(
            self.windows_directory,
            system_directory,
            "WindowsPowerShell",
            "v1.0",
            "powershell.exe",
        )


def describe_powershell_startup_failure(
    process,
    stderr_reader,
    stderr_queue,
    ready_error: Optional[str],
    *,
    max_stderr_chars: int = 2000,
) -> str:
    """Stop a failed PowerShell startup and return bounded diagnostics."""

    exit_code = process.poll() if process is not None else None
    if process is not None and exit_code is None:
        try:
            process.kill()
            exit_code = process.wait(timeout=1)
        except Exception:
            exit_code = process.poll()
    if stderr_reader is not None:
        try:
            stderr_reader.join(timeout=0.5)
        except Exception:
            pass

    stderr_lines = []
    if stderr_queue is not None:
        while True:
            try:
                line = stderr_queue.get_nowait().strip()
            except queue.Empty:
                break
            if line:
                stderr_lines.append(line)

    details = []
    if ready_error and ready_error != "未知错误":
        details.append(ready_error)
    if exit_code is not None:
        details.append(f"exit code {exit_code}")
    if stderr_lines:
        stderr = " | ".join(stderr_lines)
        if len(stderr) > max_stderr_chars:
            stderr = stderr[:max_stderr_chars] + "…"
        details.append(f"stderr: {stderr}")
    return "; ".join(details) or "未知错误"


def resolve_com_runtime(
    progid: str,
    *,
    environment: Optional[WindowsComEnvironment] = None,
) -> ComRuntimeResolution:
    """Choose a PowerShell host whose registry view can activate ``progid``."""

    environment = environment or NativeWindowsComEnvironment()
    view_bits = (32, 64) if environment.os_bitness == 64 else (32,)
    observed_registrations = []
    for view_bitness in view_bits:
        try:
            registration = environment.inspect_registration(progid, view_bitness)
        except Exception as exc:
            registration = ComRegistration(
                progid=progid,
                view_bitness=view_bitness,
                problem=f"registry inspection failed: {type(exc).__name__}: {exc}",
            )
        observed_registrations.append(registration)
    registrations = tuple(observed_registrations)
    current = next(
        (
            registration
            for registration in registrations
            if registration.view_bitness == environment.process_bitness
            and registration.activatable
        ),
        None,
    )
    selected = current or next(
        (registration for registration in registrations if registration.activatable),
        None,
    )
    if selected is None:
        return ComRuntimeResolution(
            progid=progid,
            process_bitness=environment.process_bitness,
            os_bitness=environment.os_bitness,
            registrations=registrations,
        )
    return ComRuntimeResolution(
        progid=progid,
        process_bitness=environment.process_bitness,
        os_bitness=environment.os_bitness,
        registrations=registrations,
        selected_view_bitness=selected.view_bitness,
        powershell_executable=environment.powershell_executable(selected.view_bitness),
    )
