"""Measure real occupied workbooks through the indexed calculation/report path.

Artifacts contain workbook evidence and must stay in validation_screenshots.
Each repetition is an isolated process; a budget refusal fails the capacity gate.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import statistics
import subprocess
import sys
import time
from pathlib import Path


def implementation_fingerprint() -> dict:
    """Bind measurements to source bytes even in a dirty development checkout."""
    from importlib.metadata import version

    import linexcel

    package = Path(linexcel.__file__).parent
    paths = sorted(package.rglob("*.py")) + sorted((package / "assets").rglob("*"))
    digest = hashlib.sha256()
    for path in paths:
        if path.is_file():
            digest.update(path.relative_to(package).as_posix().encode() + b"\0")
            digest.update(path.read_bytes())
    return {
        "packageSha256": digest.hexdigest(),
        "benchmarkSha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
        "formualizerVersion": version("formualizer"),
    }


def generate(path: Path, cells: int):
    from openpyxl import Workbook

    if cells < 10 or cells % 10:
        raise ValueError("Cell count must be a positive multiple of ten")
    book = Workbook(write_only=True)
    sheet = book.create_sheet("Data")
    for row in range(1, cells // 10 + 1):
        sheet.append([row + col for col in range(1, 10)] + [f"=SUM(A{row}:I{row})"])
    book.save(path)


def worker(root: Path, cells: int):
    implementation = implementation_fingerprint()
    if sys.platform != "win32":
        import resource

        resource.setrlimit(resource.RLIMIT_AS, (2 * 1024**3,) * 2)
    from linexcel.lazy import evaluate_node
    from linexcel.lazy_store import (
        WorkbookIndex,
        build_index,
        graph_neighborhood,
        graph_page,
    )
    from linexcel.web import TaskStore, WebConfig
    from linexcel.web_report import prepare_report, write_report

    source = root.parent / f"{cells}.xlsx"
    root.mkdir()
    index = root / "index.sqlite"
    start = time.perf_counter()
    build_index(source, index)
    imported = time.perf_counter() - start
    with WorkbookIndex(index) as store:
        occupied = store.db.execute(
            "SELECT count(*) FROM nodes WHERE kind IN ('input','cell')"
        ).fetchone()[0]
        formulas = store.db.execute(
            "SELECT count(*) FROM nodes WHERE formula IS NOT NULL"
        ).fetchone()[0]
        nodes = len(store)
    timings = []
    for number in range(20):
        row = 1 + number * (cells // 10 - 1) // 19
        node_id = f"Data!J{row}"
        start = time.perf_counter()
        page = graph_page(index, query=node_id)
        neighborhood = graph_neighborhood(index, node_id)
        timings.append(time.perf_counter() - start)
        assert page["nodes"][0]["id"] == node_id
        assert len(neighborhood["nodes"]) <= 200
    selected = f"Data!J{cells // 20}"
    start = time.perf_counter()
    result = evaluate_node(source, selected, index_path=index)
    calculated = time.perf_counter() - start
    assert result["status"] == "completed", result
    assert result["value"] == (cells // 20) * 9 + 45
    # Exercise the real frozen-report producer on the indexed large workbook.
    store = TaskStore(WebConfig(root / "workspace", retention_days=None))
    try:
        project = store.root / "benchmark" / "projects" / ("a" * 32)
        project.mkdir(parents=True)
        (project / "references").mkdir()
        import shutil

        shutil.copyfile(source, project / "workbook.xlsx")
        shutil.copyfile(index, project / "index.sqlite")
        metadata = {
            "id": "a" * 32,
            "name": "Capacity",
            "filename": source.name,
            "revision": store._revision(project),
            "createdAt": time.time(),
        }
        (project / "project.json").write_text(json.dumps(metadata))
        task = project / "tasks" / ("b" * 32)
        task.mkdir(parents=True)
        start = time.perf_counter()
        snapshot = prepare_report(store, "benchmark", "a" * 32, {"nodeIds": [selected]})
        report = write_report(snapshot, task)
        exported = time.perf_counter() - start
    finally:
        store.close()
    if sys.platform == "win32":
        from benchmark_lot2_dense_memory import get_process_working_set_bytes

        peak = get_process_working_set_bytes()[1]
    else:
        peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        peak *= 1 if sys.platform == "darwin" else 1024
    p95 = sorted(timings)[18]
    evidence = {
        "implementation": implementation,
        "implementationStable": implementation == implementation_fingerprint(),
        "occupiedCells": occupied,
        "formulas": formulas,
        "indexNodes": nodes,
        "importSeconds": imported,
        "lookupAndNeighborhoodP95Seconds": p95,
        "targetedCalculationSeconds": calculated,
        "exportSeconds": exported,
        "peakRssBytes": peak,
        "indexBytes": index.stat().st_size,
        "reportBytes": report["bytes"],
        "expectedValue": result["value"],
        "passed": implementation == implementation_fingerprint()
        and occupied == cells
        and formulas == cells // 10
        and imported <= 600
        and p95 <= 2
        and calculated <= 120
        and exported <= 600
        and peak <= 2 * 1024**3,
    }
    (root / "metrics.json").write_text(json.dumps(evidence, indent=2))
    return 0 if evidence["passed"] else 1


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--cells", nargs="+", type=int, default=[100000, 1000000])
    parser.add_argument("--repetitions", type=int, default=3)
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker:
        return worker(args.output, args.cells[0])
    args.output.mkdir(parents=True, exist_ok=False)
    report = {
        "implementation": implementation_fingerprint(),
        "machine": platform.platform(),
        "python": sys.version,
        "cpuCount": os.cpu_count(),
        "scope": "index/calculation/report; browser separately",
        "repetitions": [],
        "passed": True,
    }
    for cells in args.cells:
        source = args.output / f"{cells}.xlsx"
        generate(source, cells)
        source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
        for repetition in range(args.repetitions):
            root = args.output / f"{cells}-{repetition}"
            env = {**os.environ, "MALLOC_ARENA_MAX": "2"}
            with (args.output / f"{cells}-{repetition}.log").open("w") as log:
                try:
                    p = subprocess.run(
                        [
                            sys.executable,
                            str(Path(__file__).resolve()),
                            "--worker",
                            "--output",
                            str(root),
                            "--cells",
                            str(cells),
                        ],
                        stdout=log,
                        stderr=log,
                        env=env,
                        timeout=900,
                    )
                    status = p.returncode
                except subprocess.TimeoutExpired:
                    status = "timeout"
            metrics = (
                json.loads((root / "metrics.json").read_text())
                if (root / "metrics.json").exists()
                else {}
            )
            passed = status == 0 and metrics.get("passed") is True
            report["repetitions"].append(
                {
                    "cells": cells,
                    "repetition": repetition,
                    "sourceSha256": source_hash,
                    "exitStatus": status,
                    **metrics,
                    "passed": passed,
                }
            )
            report["passed"] &= passed
            (args.output / "benchmark.json").write_text(json.dumps(report, indent=2))
            print(json.dumps(report["repetitions"][-1]), flush=True)
    report["importMedianSeconds"] = {
        str(cells): statistics.median(
            [
                r["importSeconds"]
                for r in report["repetitions"]
                if r["cells"] == cells and "importSeconds" in r
            ]
        )
        for cells in args.cells
        if any(
            r["cells"] == cells and "importSeconds" in r for r in report["repetitions"]
        )
    }
    (args.output / "benchmark.json").write_text(json.dumps(report, indent=2))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
