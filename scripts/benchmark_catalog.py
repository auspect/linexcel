"""Reproducible catalog timings and label coverage; outputs stay local.

uv run python scripts/benchmark_catalog.py --workbook INPUT.xlsx --output-dir
validation_screenshots/catalog-benchmark --coverage-workbook PUBLIC.xlsx
"""

from __future__ import annotations

import argparse
import json
import statistics
import time
from pathlib import Path

from openpyxl import Workbook

from linexcel.catalog import Catalog, CatalogLimits, Source
from linexcel.lazy_store import WorkbookIndex, _unpack, build_index


def timed_search(catalog, sources):
    samples = []
    for query in ("revenue", "sales", "budget", "café", "100") * 4:
        start = time.perf_counter()
        result = catalog.search(sources, query, limit=10)
        samples.append(
            {
                "query": query,
                "ms": (time.perf_counter() - start) * 1000,
                "returned": len(result["results"]),
                "capsHit": result["coverage"]["capsHit"],
            }
        )
    timings = sorted(r["ms"] for r in samples)
    return {
        "queries": samples,
        "medianMs": statistics.median(timings),
        "p95Ms": timings[int(0.95 * (len(timings) - 1))],
        "maxMs": max(timings),
    }


def measure(workbook, directory, *, sample=2000, index_path=None):
    directory.mkdir(parents=True, exist_ok=True)
    path = index_path or directory / "index.sqlite"
    reused = path.exists()
    started = time.perf_counter()
    if not path.exists():
        build_index(workbook, path)
    indexing = time.perf_counter() - started
    source = Source("1" * 32, workbook.name, path)
    catalog = Catalog(directory / "catalog", CatalogLimits())
    started = time.perf_counter()
    catalog.ensure([source], force=True)
    build_ms = (time.perf_counter() - started) * 1000
    status = catalog.status()
    with WorkbookIndex(path) as index:
        stride = max(1, len(index) // sample)
        numeric = labelled = table = above = left = 0
        by_sheet = {}
        for sheet, row, col, blob in index.db.execute(
            "SELECT sheet,row,col,data FROM nodes WHERE seq % ?=0 "
            "AND kind IN ('cell','input','array_result') ORDER BY seq",
            (stride,),
        ):
            node = _unpack(blob)
            if type(node.get("value")) not in {int, float}:
                continue
            labels = index.cell_labels(sheet, row, col)
            numeric += 1
            labelled += bool(labels)
            table += any(label["method"] == "table_column" for label in labels)
            above += any(label["method"] == "above_text" for label in labels)
            left += any(label["method"] == "left_text" for label in labels)
            counts = by_sheet.setdefault(sheet, {"numeric": 0, "labelled": 0})
            counts["numeric"] += 1
            counts["labelled"] += bool(labels)
    return {
        "file": workbook.name,
        "workbookBytes": workbook.stat().st_size,
        "indexBytes": path.stat().st_size,
        "indexSeconds": indexing,
        "indexReused": reused,
        "catalogBuildMs": build_ms,
        "catalogBytes": status["bytes"],
        "catalogVsIndex": status["bytes"] / path.stat().st_size,
        "coverage": status["coverage"],
        "latency": timed_search(catalog, [source]),
        "labelSample": {
            "method": "every stride-th indexed node, then numeric cells",
            "stride": stride,
            "numeric": numeric,
            "labelled": labelled,
            "unlabelled": numeric - labelled,
            "table": table,
            "above": above,
            "left": left,
            "bySheet": by_sheet,
        },
    }


def small_library(directory):
    directory.mkdir(parents=True, exist_ok=True)
    sources = []
    for number in range(1, 31):
        book = Workbook()
        sheet = book.active
        sheet.append(["Région", "Revenue café", "Budget"])
        for row in range(2, 22):
            sheet.append([f"Region {row}", number * row, f"=B{row}*2"])
        path = directory / f"workbook{number}.xlsx"
        if not path.exists():
            book.save(path)
        index = directory / f"index{number}.sqlite"
        if not index.exists():
            build_index(path, index)
        sources.append(Source(f"{number:032x}", path.name, index))
    catalog = Catalog(directory / "catalog")
    started = time.perf_counter()
    catalog.ensure(sources, force=True)
    build_ms = (time.perf_counter() - started) * 1000
    status = catalog.status()
    return {
        "workbooks": len(sources),
        "indexBytes": sum(s.path.stat().st_size for s in sources),
        "catalogBytes": status["bytes"],
        "catalogBuildMs": build_ms,
        "coverage": status["coverage"],
        "latency": timed_search(catalog, sources),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--coverage-workbook", type=Path, action="append", default=[])
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--reuse-large-index", type=Path)
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    result = {
        "large": measure(
            args.workbook, args.output_dir / "large", index_path=args.reuse_large_index
        )
    }
    target = args.output_dir / "measurements.json"
    target.write_text(json.dumps(result, ensure_ascii=False, indent=2), "utf-8")
    result["smallLibrary"] = small_library(args.output_dir / "small")
    target.write_text(json.dumps(result, ensure_ascii=False, indent=2), "utf-8")
    result["realWorkbooks"] = [
        measure(path, args.output_dir / f"real{i}")
        for i, path in enumerate(args.coverage_workbook)
    ]
    target.write_text(json.dumps(result, ensure_ascii=False, indent=2), "utf-8")
    print(target)


if __name__ == "__main__":
    main()
