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


class _CompletionPort(ctypes.Structure):
    _fields_ = [("key", ctypes.c_void_p), ("port", wintypes.HANDLE)]


JOB_OBJECT_MSG_JOB_MEMORY_LIMIT = 10


class WorkerJob:
    """Caps the committed memory of a worker and all its descendants.

    After ``close``, ``memory_limit_hit`` says whether any of them was refused
    memory by the cap, and ``peak_bytes`` is the job's peak commit. Windows
    counts a refused allocation in that peak, so it can exceed the cap.
    """

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
        api.CreateIoCompletionPort.argtypes = [
            wintypes.HANDLE,
            wintypes.HANDLE,
            ctypes.c_size_t,
            wintypes.DWORD,
        ]
        api.CreateIoCompletionPort.restype = wintypes.HANDLE
        api.GetQueuedCompletionStatus.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(wintypes.DWORD),
            ctypes.POINTER(ctypes.c_size_t),
            ctypes.POINTER(ctypes.c_void_p),
            wintypes.DWORD,
        ]
        api.GetQueuedCompletionStatus.restype = wintypes.BOOL
        self.api = api
        self.port = None
        self.memory_limit_hit = False
        self.peak_bytes: int | None = None
        self.handle = api.CreateJobObjectW(None, None)
        if not self.handle:
            raise ctypes.WinError(ctypes.get_last_error())
        limits = _Extended()
        limits.basic.flags = 0x2000 | 0x200  # KILL_ON_JOB_CLOSE | JOB_MEMORY
        limits.job_memory = memory_mb * 1024 * 1024
        try:
            # The job posts JOB_OBJECT_MSG_JOB_MEMORY_LIMIT to this port when
            # any process in it is refused memory by the cap.
            self.port = api.CreateIoCompletionPort(wintypes.HANDLE(-1), None, 0, 1)
            if not self.port:
                raise ctypes.WinError(ctypes.get_last_error())
            association = _CompletionPort(None, self.port)
            if not api.SetInformationJobObject(
                self.handle, 7, ctypes.byref(association), ctypes.sizeof(association)
            ):
                raise ctypes.WinError(ctypes.get_last_error())
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
                self._read_memory()
            finally:
                self.api.CloseHandle(self.handle)
                self.handle = None
                if self.port:
                    self.api.CloseHandle(self.port)
                    self.port = None

    def _read_memory(self):
        message = wintypes.DWORD()
        key = ctypes.c_size_t()
        overlapped = ctypes.c_void_p()
        while self.api.GetQueuedCompletionStatus(
            self.port,
            ctypes.byref(message),
            ctypes.byref(key),
            ctypes.byref(overlapped),
            0,
        ):
            if message.value == JOB_OBJECT_MSG_JOB_MEMORY_LIMIT:
                self.memory_limit_hit = True
        info = _Extended()
        if self.api.QueryInformationJobObject(
            self.handle, 9, ctypes.byref(info), ctypes.sizeof(info), None
        ):
            self.peak_bytes = info.peak_job

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
