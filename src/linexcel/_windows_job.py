"""Windows Job Object: hard aggregate memory cap and descendant cleanup."""

from __future__ import annotations

import ctypes
import sys
import time
from ctypes import wintypes
from typing import Any


class _Basic(ctypes.Structure):
    _fields_ = [
        ("user", ctypes.c_int64),
        ("job_user", ctypes.c_int64),
        ("flags", wintypes.DWORD),
        ("min_working", ctypes.c_size_t),
        ("max_working", ctypes.c_size_t),
        ("active", wintypes.DWORD),
        ("affinity", ctypes.c_size_t),
        ("priority", wintypes.DWORD),
        ("scheduling", wintypes.DWORD),
    ]


class _IO(ctypes.Structure):
    _fields_ = [
        (name, ctypes.c_uint64)
        for name in (
            "read_ops",
            "write_ops",
            "other_ops",
            "read_bytes",
            "write_bytes",
            "other_bytes",
        )
    ]


class _Extended(ctypes.Structure):
    _fields_ = [
        ("basic", _Basic),
        ("io", _IO),
        ("process_memory", ctypes.c_size_t),
        ("job_memory", ctypes.c_size_t),
        ("peak_process", ctypes.c_size_t),
        ("peak_job", ctypes.c_size_t),
    ]


class _Accounting(ctypes.Structure):
    _fields_ = [
        (name, ctypes.c_int64)
        for name in ("user", "kernel", "period_user", "period_kernel")
    ] + [
        (name, wintypes.DWORD)
        for name in (
            "faults",
            "total_processes",
            "active_processes",
            "terminated_processes",
        )
    ]


class WorkerJob:
    api: Any

    def __init__(self, process, memory_mb: int):
        if sys.platform != "win32":
            raise RuntimeError("Windows Job Objects require Windows")
        api = ctypes.WinDLL("kernel32", use_last_error=True)
        api.CreateJobObjectW.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR]
        api.CreateJobObjectW.restype = wintypes.HANDLE
        api.SetInformationJobObject.argtypes = [
            wintypes.HANDLE,
            ctypes.c_int,
            ctypes.c_void_p,
            wintypes.DWORD,
        ]
        api.SetInformationJobObject.restype = wintypes.BOOL
        api.AssignProcessToJobObject.argtypes = [wintypes.HANDLE, wintypes.HANDLE]
        api.AssignProcessToJobObject.restype = wintypes.BOOL
        api.CloseHandle.argtypes = [wintypes.HANDLE]
        api.CloseHandle.restype = wintypes.BOOL
        api.TerminateJobObject.argtypes = [wintypes.HANDLE, wintypes.UINT]
        api.TerminateJobObject.restype = wintypes.BOOL
        api.QueryInformationJobObject.argtypes = [
            wintypes.HANDLE,
            ctypes.c_int,
            ctypes.c_void_p,
            wintypes.DWORD,
            ctypes.c_void_p,
        ]
        api.QueryInformationJobObject.restype = wintypes.BOOL
        self.api = api
        self.handle = api.CreateJobObjectW(None, None)
        if not self.handle:
            raise ctypes.WinError(ctypes.get_last_error())
        limits = _Extended()
        limits.basic.flags = 0x2000 | 0x200  # KILL_ON_JOB_CLOSE | JOB_MEMORY
        limits.job_memory = memory_mb * 1024 * 1024
        try:
            if not api.SetInformationJobObject(
                self.handle, 9, ctypes.byref(limits), ctypes.sizeof(limits)
            ):
                raise ctypes.WinError(ctypes.get_last_error())
            if not api.AssignProcessToJobObject(
                self.handle, wintypes.HANDLE(int(process._handle))
            ):
                raise ctypes.WinError(ctypes.get_last_error())
        except BaseException:
            self.close()
            raise

    def close(self):
        if self.handle:
            try:
                if not self.api.TerminateJobObject(self.handle, 1):
                    raise OSError("Could not terminate isolated worker Job")
                deadline = time.monotonic() + 2
                while True:
                    accounting = _Accounting()
                    if not self.api.QueryInformationJobObject(
                        self.handle,
                        1,
                        ctypes.byref(accounting),
                        ctypes.sizeof(accounting),
                        None,
                    ):
                        raise OSError("Could not verify isolated worker cleanup")
                    if not accounting.active_processes:
                        break
                    if time.monotonic() >= deadline:
                        raise TimeoutError(
                            "Isolated descendants did not stop within 2 seconds"
                        )
                    time.sleep(0.005)
            finally:
                self.api.CloseHandle(self.handle)
                self.handle = None

    def resume(self, process):
        if sys.platform != "win32":
            raise RuntimeError("Windows process resumption requires Windows")
        # Popen closes the primary-thread handle. Resume the suspended process
        # only after its Job is installed; no workbook/Python startup runs first.
        native = ctypes.WinDLL("ntdll")
        native.NtResumeProcess.argtypes = [wintypes.HANDLE]
        native.NtResumeProcess.restype = ctypes.c_long
        status = native.NtResumeProcess(wintypes.HANDLE(int(process._handle)))
        if status < 0:
            raise OSError(f"Could not resume isolated worker (NTSTATUS {status:#x})")
