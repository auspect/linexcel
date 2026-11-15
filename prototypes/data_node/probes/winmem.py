"""Peak memory of the current process on Windows (no psutil)."""

import ctypes
import os
import sys
from ctypes import wintypes


class _PMC(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD),
        ("PageFaultCount", wintypes.DWORD),
        ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t),
        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
        ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t),
    ]


def peak_mb() -> dict:
    if os.name != "nt":
        import resource

        return {"peakRssMB": resource.getrusage(resource.RUSAGE_SELF).ru_maxrss // 1024}
    pmc = _PMC()
    pmc.cb = ctypes.sizeof(_PMC)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.GetCurrentProcess.restype = wintypes.HANDLE
    kernel32.K32GetProcessMemoryInfo.argtypes = [
        wintypes.HANDLE,
        ctypes.POINTER(_PMC),
        wintypes.DWORD,
    ]
    kernel32.K32GetProcessMemoryInfo.restype = wintypes.BOOL
    if not kernel32.K32GetProcessMemoryInfo(
        kernel32.GetCurrentProcess(), ctypes.byref(pmc), pmc.cb
    ):
        raise OSError(ctypes.get_last_error(), "GetProcessMemoryInfo failed")
    return {
        "peakWorkingSetMB": round(pmc.PeakWorkingSetSize / 1e6, 1),
        "peakCommitMB": round(pmc.PeakPagefileUsage / 1e6, 1),
    }


if __name__ == "__main__":
    print(peak_mb(), file=sys.stderr)
