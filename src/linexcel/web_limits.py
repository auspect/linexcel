"""Resource limits of the web application: defaults and failure wording.

The launcher derives its memory and storage budgets from the machine. An
embedding host passes its own to ``create_app``; ``WebConfig`` keeps fixed
defaults for it.
"""

from __future__ import annotations

import ctypes
import os
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from linexcel.web import WebConfig

MiB = 1048576
GiB = 1024 * MiB
MAX_TASK_SECONDS = 3600
# Whole-workbook work. AI documentation keeps the long class it shared with
# imports; the page sends its own AI duration anyway.
LONG_OPERATIONS = frozenset(
    {"import", "build_graph", "export_html", "capture", "document", "describe_capture"}
)
LIMIT_KINDS = frozenset({"timeout", "memory_limit", "storage_limit"})


def physical_memory_mb() -> int | None:
    try:
        if sys.platform == "win32":

            class Status(ctypes.Structure):
                _fields_ = [("length", ctypes.c_uint32), ("load", ctypes.c_uint32)] + [
                    (name, ctypes.c_uint64)
                    for name in (
                        "total_physical",
                        "available_physical",
                        "total_page_file",
                        "available_page_file",
                        "total_virtual",
                        "available_virtual",
                        "available_extended_virtual",
                    )
                ]

            status = Status(length=ctypes.sizeof(Status))
            if not ctypes.WinDLL("kernel32").GlobalMemoryStatusEx(ctypes.byref(status)):
                return None
            return status.total_physical // MiB
        return os.sysconf("SC_PAGE_SIZE") * os.sysconf("SC_PHYS_PAGES") // MiB
    except (AttributeError, OSError, ValueError):
        return None


def launcher_memory_mb(physical_mb: int | None) -> int:
    """Half of the machine for workers, never less than the former 2 GiB."""
    return max(2048, (physical_mb or 0) // 2)


def launcher_storage_mb(used_bytes: int, free_bytes: int, total_bytes: int) -> int:
    """What the data directory holds, plus its disk's free space minus a reserve.

    The reserve is a tenth of the disk, at most 10 GiB. A nearly full disk still
    gives the former 2 GiB of room when it has that much free.
    """
    reserve = min(total_bytes // 10, 10 * GiB)
    room = max(free_bytes - reserve, min(free_bytes, 2 * GiB))
    return max(1, (used_bytes + room) // MiB)


def directory_bytes(root: Path) -> int:
    total = 0
    for path in root.rglob("*"):
        try:
            if path.is_file():
                total += path.stat().st_size
        except OSError:
            pass  # A file can be replaced or removed while it is counted.
    return total


@dataclass(frozen=True)
class LauncherLimits:
    memory_mb: int
    storage_mb: int
    storage_used_bytes: int


def launcher_limits(
    data_dir: Path, memory_mb: int | None, storage_mb: int | None
) -> LauncherLimits:
    """Fill the limits the operator did not set from this machine."""
    if memory_mb is None:
        memory_mb = launcher_memory_mb(physical_memory_mb())
    root = Path(data_dir).resolve()
    used = directory_bytes(root) if root.is_dir() else 0
    if storage_mb is None:
        disk = root
        while not disk.exists() and disk != disk.parent:
            disk = disk.parent
        usage = shutil.disk_usage(disk)
        storage_mb = launcher_storage_mb(used, usage.free, usage.total)
    return LauncherLimits(memory_mb, storage_mb, used)


@dataclass(frozen=True)
class TaskDefaults:
    """Budget of an operation submitted without its own duration or memory."""

    memory_mb: int
    long_seconds: int = MAX_TASK_SECONDS
    seconds: int = 120

    @classmethod
    def for_server(cls, total_memory_mb: int, max_workers: int) -> TaskDefaults:
        share = max(1024, total_memory_mb // max_workers)
        return cls(memory_mb=min(total_memory_mb, share))

    def seconds_for(self, operation: str) -> int:
        return self.long_seconds if operation in LONG_OPERATIONS else self.seconds


def _duration_setting(operation: str | None) -> str:
    if operation in {"document", "describe_capture"}:
        return "Durée maximale (s)"
    if operation in LONG_OPERATIONS:
        return "Import, graphe, captures et rapport (secondes)"
    return "Recalcul ciblé (secondes)"


def limit_message(
    kind: str,
    config: WebConfig,
    *,
    operation: str | None = None,
    seconds: float | None = None,
    memory_mb: int | None = None,
    used_bytes: int = 0,
) -> str:
    """French source text naming the limit, its value and the remedy."""
    kept = "Le classeur et les autres résultats restent accessibles."
    if kind == "timeout":
        return (
            f"Durée maximale de {seconds:g} s atteinte. Augmentez "
            f"« {_duration_setting(operation)} » (au plus {MAX_TASK_SECONDS} s), "
            f"puis relancez. {kept}"
        )
    if kind == "memory_limit":
        return (
            f"Budget mémoire de {memory_mb} Mio atteint. Augmentez « Mémoire (MiB) » "
            f"jusqu’au plafond serveur de {config.total_memory_mb} Mio, ou relancez "
            f"le serveur avec --memory-mb pour relever ce plafond. {kept}"
        )
    if kind == "storage_limit":
        return (
            f"Quota de stockage de {config.max_storage_mb} Mio dépassé "
            f"({-(-used_bytes // MiB)} Mio requis). Supprimez des espaces de travail "
            "ou relancez le serveur avec --max-storage-mb. "
            "Les classeurs et les autres résultats restent accessibles."
        )
    raise ValueError(f"Unknown limit kind: {kind}")
