#!/usr/bin/env python3
"""Benchmark 2B: Recovery cost on broken references (Issue #65, 2B).

Protocol requirements from Issue #65 & PLAN_DELEGATION.md:
1. Extend dense_chain_broken_ref across progressive sizes (e.g. 50, 150, 300, 600) with a valid control of identical structure.
2. Distinct variants:
   - valid_control (no broken ref)
   - broken_sheet_ref (e.g. =Ghost!A1)
   - broken_ref_literal (e.g. #REF!)
   - missing_name (e.g. =MissingRate)
   - unresolved_external (e.g. =[External.xlsx]Sheet1!A1)
   - fan_out (1 broken ref read by many independent formulas)
3. Separate phase timings: import, native evaluation, recovery, decomposition.
4. Measure peak RSS, fallback evaluations, budget consumption, warnings.
5. Fresh subprocess per scenario run for isolated memory tracking.
"""

from __future__ import annotations

import argparse
import json
import platform
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import openpyxl  # noqa: E402

from linexcel import analyze  # noqa: E402


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
        peak = usage.ru_maxrss * 1024 if sys.platform != "darwin" else usage.ru_maxrss
        return peak, peak


def build_scenario_workbook(variant: str, size: int, out_path: Path) -> dict[str, Any]:
    """Build a workbook with the specified size and broken reference variant."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"

    if variant == "valid_control":
        ws["Z1"] = 1
        for r in range(1, size + 1):
            ws.cell(row=r, column=1, value=f"=SUM(A1:A{max(1, r - 1)}) + Z1")
    elif variant == "broken_sheet_ref":
        ws["Z1"] = "=Ghost!A1"
        for r in range(1, size + 1):
            ws.cell(row=r, column=1, value=f"=SUM(A1:A{max(1, r - 1)}) + Z1")
    elif variant == "broken_ref_literal":
        ws["Z1"] = "=#REF!"
        for r in range(1, size + 1):
            ws.cell(row=r, column=1, value=f"=SUM(A1:A{max(1, r - 1)}) + Z1")
    elif variant == "missing_name":
        for r in range(1, size + 1):
            ws.cell(row=r, column=1, value=f"=SUM(A1:A{max(1, r - 1)}) + MissingRate")
    elif variant == "unresolved_external":
        ws["Z1"] = "='[MissingExternal.xlsx]Data'!A1"
        for r in range(1, size + 1):
            ws.cell(row=r, column=1, value=f"=SUM(A1:A{max(1, r - 1)}) + Z1")
    elif variant == "fan_out_broken":
        ws["Z1"] = "=Ghost!A1"
        for r in range(1, size + 1):
            # Independent formulas each reading Z1 (fan-out)
            ws.cell(row=r, column=1, value=f"=Z1 * {r}")
    else:
        raise ValueError(f"Unknown variant: {variant}")

    wb.save(out_path)
    return {
        "variant": variant,
        "size": size,
        "formulas_count": size
        + (
            1
            if variant
            in {
                "broken_sheet_ref",
                "broken_ref_literal",
                "unresolved_external",
                "fan_out_broken",
            }
            else 0
        ),
        "file_size_bytes": out_path.stat().st_size,
    }


def _worker_analyze(file_path: str) -> None:
    """Analyze workbook in an isolated subprocess to capture accurate peak working set."""
    ws_initial, peak_initial = get_process_working_set_bytes()
    data = Path(file_path).read_bytes()

    t0 = time.perf_counter()
    result = analyze(data, filename="benchmark.xlsx")
    total_elapsed = time.perf_counter() - t0

    ws_final, peak_final = get_process_working_set_bytes()
    meta = result.graph.get("meta", {})

    metrics = {
        "elapsed_seconds": round(total_elapsed, 4),
        "peak_rss_mb": round(peak_final / (1024 * 1024), 2),
        "rss_delta_mb": round((peak_final - ws_initial) / (1024 * 1024), 2),
        "total_nodes": result.stats.get("totalNodes", 0),
        "total_edges": result.stats.get("totalEdges", 0),
        "total_formulas": result.stats.get("totalFormulas", 0),
        "quarantined_count": len(meta.get("quarantined", {})),
        "warnings_count": len(result.warnings),
        "execution_status": meta.get("execution", {}).get("status", "unknown"),
        "phase_metrics": meta.get("execution", {}).get("phaseMetrics", {}),
    }
    print(json.dumps(metrics))


def run_single_analysis(file_path: Path) -> dict[str, Any]:
    cmd = [
        sys.executable,
        "-c",
        (
            "from scripts.benchmark_lot2_broken_refs import _worker_analyze; "
            f"_worker_analyze(r'{file_path}')"
        ),
    ]
    proc = subprocess.run(
        cmd,
        cwd=str(ROOT),
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"Subprocess failed:\nSTDOUT: {proc.stdout}\nSTDERR: {proc.stderr}"
        )
    return json.loads(proc.stdout.strip())


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Benchmark recovery cost on broken references (Issue #65, 2B)"
    )
    parser.add_argument(
        "--sizes",
        type=str,
        default="50,150,300,600",
        help="Comma-separated chain lengths",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=ROOT
        / "validation_screenshots/delegation-20260920/benchmark_2b_broken_refs.json",
    )
    args = parser.parse_args()

    sizes = [int(s.strip()) for s in args.sizes.split(",") if s.strip()]
    variants = [
        "valid_control",
        "broken_sheet_ref",
        "broken_ref_literal",
        "missing_name",
        "unresolved_external",
        "fan_out_broken",
    ]

    print("=" * 80)
    print("Benchmark 2B: Recovery Cost on Broken References")
    print(f"Sizes tested: {sizes}")
    print(f"Variants    : {variants}")
    print(
        f"Platform    : {platform.system()} {platform.release()} ({platform.machine()})"
    )
    print("=" * 80)

    report: dict[str, Any] = {
        "timestamp_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
            "python": platform.python_version(),
        },
        "sizes": sizes,
        "results": {},
    }

    with tempfile.TemporaryDirectory(prefix="linexcel-bench-2b-") as temp_dir:
        temp_root = Path(temp_dir)

        for variant in variants:
            print(f"\n--- Variant: {variant} ---")
            report["results"][variant] = {}

            for size in sizes:
                wb_path = temp_root / f"wb_{variant}_{size}.xlsx"
                meta = build_scenario_workbook(variant, size, wb_path)
                res = run_single_analysis(wb_path)

                print(
                    f"  Size {size:4d}: {res['elapsed_seconds']:7.4f}s | "
                    f"peak RSS: {res['peak_rss_mb']:6.1f} MB | "
                    f"nodes: {res['total_nodes']:4d} | "
                    f"edges: {res['total_edges']:4d} | "
                    f"quarantined: {res['quarantined_count']:2d}"
                )
                report["results"][variant][str(size)] = {
                    "metadata": meta,
                    "metrics": res,
                }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print("\n" + "=" * 80)
    print(f"Results successfully saved to {args.output}")
    print("=" * 80)
    return 0


if __name__ == "__main__":
    sys.exit(main())
