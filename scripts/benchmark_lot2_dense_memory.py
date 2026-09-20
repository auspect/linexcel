#!/usr/bin/env python3
"""Benchmark 2A: Dense reading memory RSS across declared cell thresholds (1M, 5M, 10M, 19M).

Protocol requirements from Issue #65 & PLAN_DELEGATION.md:
1. Synthetic workbooks at 1M, 5M, 10M, 19M declared cells with shapes and occupancy rates.
2. Separate subprocess for each repetition to measure true process peak working set (RSS).
3. Measure load_cached_values in isolation (excluding workbook generation time).
4. Measure end-to-end time separately.
5. 3 repetitions per tier, reporting median, min, max, stddev.
6. Compare python-calamine (fast dense reader) vs openpyxl (fallback streaming reader).
"""

from __future__ import annotations

import argparse
import json
import platform
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import openpyxl  # noqa: E402

from linexcel.loader import declared_cells, load_cached_values  # noqa: E402


def get_process_working_set_bytes() -> tuple[int, int]:
    """Return (current_working_set_bytes, peak_working_set_bytes) of the current process."""
    if sys.platform == "win32":
        import ctypes.wintypes

        class PMC(ctypes.Structure):
            _fields_ = [
                ("cb", ctypes.wintypes.DWORD),
                ("PageFaultCount", ctypes.wintypes.DWORD),
                ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t),
                ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t),
            ]

        p = PMC()
        p.cb = ctypes.sizeof(PMC)
        fn = ctypes.windll.psapi.GetProcessMemoryInfo
        fn.argtypes = [
            ctypes.wintypes.HANDLE,
            ctypes.POINTER(PMC),
            ctypes.wintypes.DWORD,
        ]
        fn.restype = ctypes.wintypes.BOOL
        ok = fn(ctypes.windll.kernel32.GetCurrentProcess(), ctypes.byref(p), p.cb)
        if not ok:
            return 0, 0
        return int(p.WorkingSetSize), int(p.PeakWorkingSetSize)
    else:
        import resource

        usage = resource.getrusage(resource.RUSAGE_SELF)
        # On Linux ru_maxrss is in KiB, on macOS in bytes
        peak = usage.ru_maxrss * 1024 if sys.platform != "darwin" else usage.ru_maxrss
        return peak, peak


def build_synthetic_tier(
    rows: int, cols: int, populated_rows: int, out_path: Path
) -> dict[str, Any]:
    """Build a synthetic workbook with specific rows x cols and populated rows."""
    t0 = time.perf_counter()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"

    # Fill the first `populated_rows` with numeric values
    cells_populated = 0
    for r in range(1, min(rows, populated_rows) + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c, value=r * 1000 + c)
            cells_populated += 1

    # Set the corner cell to establish the full declared dimension
    if rows > populated_rows:
        ws.cell(row=rows, column=cols, value=999999)
        cells_populated += 1

    wb.save(out_path)
    build_duration = time.perf_counter() - t0
    file_size_bytes = out_path.stat().st_size

    data = out_path.read_bytes()
    declared = declared_cells(data)

    return {
        "rows": rows,
        "cols": cols,
        "declared_cells": declared,
        "populated_cells": cells_populated,
        "occupancy_rate": cells_populated / declared if declared > 0 else 0.0,
        "build_seconds": round(build_duration, 4),
        "file_size_bytes": file_size_bytes,
        "file_size_mb": round(file_size_bytes / (1024 * 1024), 3),
    }


def _worker_measure(file_path: str, max_dense_cells: int | None) -> None:
    """Entry point executed in an isolated child process to measure memory and time."""
    ws_initial, peak_initial = get_process_working_set_bytes()
    data = Path(file_path).read_bytes()
    ws_after_read, peak_after_read = get_process_working_set_bytes()

    warnings: list[str] = []
    t0 = time.perf_counter()
    cached = load_cached_values(
        data,
        warnings=warnings,
        max_dense_cells=max_dense_cells,
    )
    elapsed = time.perf_counter() - t0
    ws_final, peak_final = get_process_working_set_bytes()

    # Determine reader used
    reader = (
        "openpyxl"
        if any("openpyxl" in w.lower() for w in warnings) or max_dense_cells == 0
        else "calamine"
    )

    result = {
        "elapsed_seconds": round(elapsed, 4),
        "cells_retained": len(cached),
        "reader": reader,
        "warnings": warnings,
        "ws_initial_mb": round(ws_initial / (1024 * 1024), 2),
        "peak_rss_mb": round(peak_final / (1024 * 1024), 2),
        "rss_delta_mb": round((peak_final - ws_initial) / (1024 * 1024), 2),
    }
    print(json.dumps(result))


def run_single_measurement(
    file_path: Path, max_dense_cells: int | None
) -> dict[str, Any]:
    """Launch a fresh python subprocess to measure working set without residual heap."""
    cmd = [
        sys.executable,
        "-c",
        f"from scripts.benchmark_lot2_dense_memory import _worker_measure; _worker_measure(r'{file_path}', {max_dense_cells})",
    ]
    proc = subprocess.run(
        cmd,
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        check=True,
    )
    return json.loads(proc.stdout.strip())


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Benchmark dense memory reading RSS (Issue #65, 2A)"
    )
    parser.add_argument(
        "--repetitions", type=int, default=3, help="Number of repetitions per tier"
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=ROOT
        / "validation_screenshots/delegation-20260920/benchmark_2a_dense_memory.json",
    )
    args = parser.parse_args()

    tiers = [
        {"name": "1M", "rows": 10_000, "cols": 100, "populated_rows": 100},
        {"name": "5M", "rows": 50_000, "cols": 100, "populated_rows": 100},
        {"name": "10M", "rows": 100_000, "cols": 100, "populated_rows": 100},
        {"name": "19M", "rows": 190_000, "cols": 100, "populated_rows": 100},
    ]

    print("=" * 80)
    print("Benchmark 2A: Dense Reading Memory RSS across Declared Tiers")
    print(f"Repetitions per tier: {args.repetitions}")
    print(f"Platform: {platform.system()} {platform.release()} ({platform.machine()})")
    print(f"Python: {platform.python_version()}")
    print("=" * 80)

    report: dict[str, Any] = {
        "timestamp_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
            "python": platform.python_version(),
        },
        "repetitions": args.repetitions,
        "tiers": {},
    }

    with tempfile.TemporaryDirectory(prefix="linexcel-bench-2a-") as temp_dir:
        temp_root = Path(temp_dir)

        for tier in tiers:
            name = tier["name"]
            wb_path = temp_root / f"synth_{name}.xlsx"
            print(
                f"\nBuilding tier {name} ({tier['rows']} rows x {tier['cols']} cols)..."
            )
            meta = build_synthetic_tier(
                tier["rows"], tier["cols"], tier["populated_rows"], wb_path
            )
            print(f"  Declared cells : {meta['declared_cells']:,}")
            print(
                f"  Populated cells: {meta['populated_cells']:,} ({meta['occupancy_rate']:.4%})"
            )
            print(
                f"  File size      : {meta['file_size_mb']} MB (built in {meta['build_seconds']} s)"
            )

            tier_results: dict[str, Any] = {"metadata": meta, "modes": {}}

            # Mode 1: Calamine (default max_dense_cells = 20_000_000)
            calamine_runs = []
            for rep in range(1, args.repetitions + 1):
                res = run_single_measurement(wb_path, max_dense_cells=20_000_000)
                calamine_runs.append(res)
                print(
                    f"  [Calamine run {rep}/{args.repetitions}] {res['elapsed_seconds']:.4f}s | peak RSS: {res['peak_rss_mb']:.1f} MB (delta: +{res['rss_delta_mb']:.1f} MB) | retained: {res['cells_retained']}"
                )

            tier_results["modes"]["calamine"] = {
                "runs": calamine_runs,
                "time_median_s": round(
                    statistics.median([r["elapsed_seconds"] for r in calamine_runs]), 4
                ),
                "time_min_s": round(
                    min(r["elapsed_seconds"] for r in calamine_runs), 4
                ),
                "time_max_s": round(
                    max(r["elapsed_seconds"] for r in calamine_runs), 4
                ),
                "peak_rss_median_mb": round(
                    statistics.median([r["peak_rss_mb"] for r in calamine_runs]), 1
                ),
                "peak_rss_min_mb": round(
                    min(r["peak_rss_mb"] for r in calamine_runs), 1
                ),
                "peak_rss_max_mb": round(
                    max(r["peak_rss_mb"] for r in calamine_runs), 1
                ),
                "rss_delta_median_mb": round(
                    statistics.median([r["rss_delta_mb"] for r in calamine_runs]), 1
                ),
            }

            # Mode 2: Openpyxl fallback (forced by max_dense_cells=0)
            openpyxl_runs = []
            for rep in range(1, args.repetitions + 1):
                res = run_single_measurement(wb_path, max_dense_cells=0)
                openpyxl_runs.append(res)
                print(
                    f"  [Openpyxl run {rep}/{args.repetitions}] {res['elapsed_seconds']:.4f}s | peak RSS: {res['peak_rss_mb']:.1f} MB (delta: +{res['rss_delta_mb']:.1f} MB) | retained: {res['cells_retained']}"
                )

            tier_results["modes"]["openpyxl"] = {
                "runs": openpyxl_runs,
                "time_median_s": round(
                    statistics.median([r["elapsed_seconds"] for r in openpyxl_runs]), 4
                ),
                "time_min_s": round(
                    min(r["elapsed_seconds"] for r in openpyxl_runs), 4
                ),
                "time_max_s": round(
                    max(r["elapsed_seconds"] for r in openpyxl_runs), 4
                ),
                "peak_rss_median_mb": round(
                    statistics.median([r["peak_rss_mb"] for r in openpyxl_runs]), 1
                ),
                "peak_rss_min_mb": round(
                    min(r["peak_rss_mb"] for r in openpyxl_runs), 1
                ),
                "peak_rss_max_mb": round(
                    max(r["peak_rss_mb"] for r in openpyxl_runs), 1
                ),
                "rss_delta_median_mb": round(
                    statistics.median([r["rss_delta_mb"] for r in openpyxl_runs]), 1
                ),
            }

            report["tiers"][name] = tier_results

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print("\n" + "=" * 80)
    print(f"Results successfully saved to {args.output}")
    print("=" * 80)
    return 0


if __name__ == "__main__":
    sys.exit(main())
