"""Worker startup and platform-specific memory supervision."""

from __future__ import annotations

import ctypes
import errno
import os
import sys
from pathlib import Path


def worker_command(module: str, root: Path, memory_mb: int) -> list[str]:
    if os.name == "nt" or sys.platform == "darwin":
        # Windows uses a Job; macOS needs parent-side RSS supervision because
        # its reserved address space can already exceed the requested budget.
        return [sys.executable, "-X", "faulthandler", "-m", module, str(root)]
    bootstrap = (
        "import resource,runpy,sys; "
        "resource.setrlimit(resource.RLIMIT_AS,"
        f"({memory_mb * 1048576},)*2); "
        f"sys.argv=[{module!r},sys.argv[1]]; "
        f"runpy.run_module({module!r},run_name='__main__')"
    )
    return [sys.executable, "-X", "faulthandler", "-c", bootstrap, str(root)]


class _RUsageInfoV0(ctypes.Structure):
    # Apple's bsd/sys/resource.h, RUSAGE_INFO_V0. Sizes are fixed on Darwin.
    _fields_ = [("uuid", ctypes.c_uint8 * 16)] + [
        (name, ctypes.c_uint64)
        for name in (
            "user_time",
            "system_time",
            "pkg_idle_wkups",
            "interrupt_wkups",
            "pageins",
            "wired_size",
            "resident_size",
            "phys_footprint",
            "proc_start_abstime",
            "proc_exit_abstime",
        )
    ]


class _DarwinGroupRSS:
    def __init__(self):
        self.lib = ctypes.CDLL("/usr/lib/libproc.dylib", use_errno=True)
        self.lib.proc_listpgrppids.argtypes = [
            ctypes.c_int,
            ctypes.c_void_p,
            ctypes.c_int,
        ]
        self.lib.proc_listpgrppids.restype = ctypes.c_int
        self.lib.proc_pid_rusage.argtypes = [
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_void_p,
        ]
        self.lib.proc_pid_rusage.restype = ctypes.c_int

    @staticmethod
    def _check_error(operation: str) -> None:
        code = ctypes.get_errno()
        if code not in (0, errno.ESRCH, errno.ENOENT):
            raise OSError(code, f"Cannot supervise macOS worker memory: {operation}")

    def __call__(self, group: int) -> int | None:
        ctypes.set_errno(0)
        count = self.lib.proc_listpgrppids(group, None, 0)
        if count <= 0:
            self._check_error("list process group")
            return None
        # libproc returns a PID count, not a byte count. Spare room allows
        # descendants to start between the sizing call and the actual read.
        pids = (ctypes.c_int * (count + 32))()
        ctypes.set_errno(0)
        count = self.lib.proc_listpgrppids(group, pids, ctypes.sizeof(pids))
        if count <= 0:
            self._check_error("read process group")
            return None
        total = 0
        observed = False
        for pid in set(pids[:count]) - {0}:
            usage = _RUsageInfoV0()
            ctypes.set_errno(0)
            if self.lib.proc_pid_rusage(pid, 0, ctypes.byref(usage)) != 0:
                self._check_error("read resident memory")
                continue  # A process can exit between enumeration and reading.
            observed = True
            total += usage.resident_size
        return total if observed else None


class MemoryMonitor:
    """macOS samples a process group's RSS; other platforms use kernel caps."""

    def __init__(self, memory_mb: int):
        self.limit = memory_mb * 1048576
        self.kind = (
            "job_commit"
            if os.name == "nt"
            else "process_group_sampled_rss"
            if sys.platform == "darwin"
            else "per_process_virtual_address"
        )
        self.peak: int | None = None
        self._reader: _DarwinGroupRSS | None = None

    def exceeded(self, group: int) -> bool:
        if self.kind != "process_group_sampled_rss":
            return False
        if self._reader is None:
            self._reader = _DarwinGroupRSS()
        resident = self._reader(group)
        if resident is None:
            return False
        self.peak = max(self.peak or 0, resident)
        return resident > self.limit

    def evidence(self) -> dict:
        result: dict = {"memoryLimitKind": self.kind}
        if self.kind == "process_group_sampled_rss":
            result["memoryPeakMiB"] = (
                round(self.peak / 1048576, 3) if self.peak is not None else None
            )
        return result
