"""Read-only, reproducible corpus probe through resource-isolated web tasks.

Outputs contain workbook data: keep them in ignored validation_screenshots.
Saved Excel caches are observational evidence, not a correctness oracle.
This supplementary probe does not replace validate_manual.py acceptance.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import math
import re
import time
from collections import Counter, defaultdict
from importlib.metadata import version
from pathlib import Path

from linexcel.web import TERMINAL, TaskStore, WebConfig


def write(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def spread(items, count):
    """Stable size-quantile selection includes both extremes."""
    if len(items) <= count:
        return items
    if count == 1:
        return [items[len(items) // 2]]
    return [items[round(i * (len(items) - 1) / (count - 1))] for i in range(count)]


def sample_nodes(nodes, count):
    candidates = [n for n in nodes if n["kind"] == "cell"]
    candidates.sort(key=lambda n: n["id"])
    # Round robin across sheets, and diverse formula signatures within sheets.
    groups = defaultdict(list)
    for node in candidates:
        formula = node.get("formula") or ""
        signature = (
            tuple(sorted(set(re.findall(r"([A-Za-z_][\w.]*)\s*\(", formula)))),
            ":" in formula,
            bool(node.get("diagnostics")),
            node.get("cache_status"),
            str(node.get("cached_value", "")).startswith("#"),
        )
        groups[(node["sheet"], signature)].append(node)
    queues = defaultdict(list)
    for (sheet, _), members in sorted(groups.items(), key=lambda item: repr(item[0])):
        queues[sheet].append(members[0])
    selected = []
    while queues and len(selected) < count:
        for sheet in sorted(list(queues)):
            selected.append(queues[sheet].pop(0))
            if not queues[sheet]:
                del queues[sheet]
            if len(selected) == count:
                break
    used = {n["id"] for n in selected}
    selected += (
        spread([n for n in candidates if n["id"] not in used], count - len(selected))
        if len(selected) < count
        else []
    )
    return selected


def wait(store, task, owner):
    while True:
        current = store.get_task(owner, task["id"])
        if current["status"] in TERMINAL:
            return current
        time.sleep(0.1)


def comparison(node, result):
    if result.get("status") != "completed":
        return "not_calculated"
    if result.get("volatile"):
        return "volatile_not_comparable"
    if node.get("cache_status") != "saved":
        return "cache_unknown"
    saved = node.get("cached_comparison_value", node.get("cached_value"))
    value = result.get("comparisonValue", result.get("value"))
    if type(value) is type(saved) and value == saved:
        return "exact_agreement"
    if (
        isinstance(saved, (int, float))
        and not isinstance(saved, bool)
        and isinstance(value, (int, float))
        and not isinstance(value, bool)
    ):
        if math.isclose(saved, value, rel_tol=1e-9, abs_tol=1e-8):
            return "numeric_tolerance_agreement"
    return "cache_difference_requires_review"


def record_exception(summary, report, stage, exc):
    report["exception"] = {
        "stage": stage,
        "type": type(exc).__name__,
        "message": str(exc),
    }
    counts = summary.setdefault("harness_errors", {})
    counts[stage] = counts.get(stage, 0) + 1


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--corpus",
        type=Path,
        required=True,
        help="Directory containing dataset directories",
    )
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="New directory below validation_screenshots",
    )
    parser.add_argument("--per-dataset", type=int, default=80)
    parser.add_argument("--nodes", type=int, default=12)
    parser.add_argument("--seconds", type=float, default=30)
    parser.add_argument("--evaluation-seconds", type=float, default=15)
    parser.add_argument("--memory-mb", type=int, default=768)
    parser.add_argument("--inventory-only", action="store_true")
    args = parser.parse_args()
    if args.per_dataset < 1 or args.nodes < 1:
        parser.error("Sample counts must be positive")
    output = args.output.resolve()
    allowed = Path(__file__).resolve().parents[1] / "validation_screenshots"
    if not output.is_relative_to(allowed.resolve()):
        parser.error("Reports must stay under ignored validation_screenshots")
    output.mkdir(parents=True, exist_ok=False)
    source_root = Path(__file__).resolve().parents[1]
    write(
        output / "environment.json",
        {
            "versions": {
                name: version(name) for name in ("linexcel", "openpyxl", "formualizer")
            },
            "source_sha256": {
                name: hashlib.sha256((source_root / name).read_bytes()).hexdigest()
                for name in (
                    "src/linexcel/lazy.py",
                    "src/linexcel/web.py",
                    "src/linexcel/web_worker.py",
                    "scripts/validate_lazy_corpus.py",
                )
            },
        },
    )
    root = args.corpus.resolve()
    inventory = []
    by_dataset = defaultdict(list)
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(root).as_posix()
        entry = {
            "path": rel,
            "dataset": rel.split("/")[0],
            "bytes": path.stat().st_size,
            "extension": path.suffix.lower(),
        }
        entry["eligible"] = entry["extension"] in {".xlsx", ".xlsm"}
        entry["exclusion"] = (
            None
            if entry["eligible"]
            else "unsupported_legacy_format"
            if entry["extension"] in {".xls", ".xlsb"}
            else "archive_or_other_file_not_imported"
        )
        inventory.append(entry)
        if entry["eligible"]:
            by_dataset[entry["dataset"]].append(entry)
    selected = []
    for dataset, entries in sorted(by_dataset.items()):
        selected.extend(
            spread(
                sorted(entries, key=lambda e: (e["bytes"], e["path"])), args.per_dataset
            )
        )
    for entry in selected:
        data = (root / entry["path"]).read_bytes()
        entry["sha256"] = hashlib.sha256(data).hexdigest()
        entry["signature"] = (
            "zip_ooxml_candidate"
            if data.startswith(b"PK")
            else "ole_compound"
            if data.startswith(bytes.fromhex("d0cf11e0a1b11ae1"))
            else "other"
        )
    hashes = defaultdict(list)
    for entry in selected:
        hashes[entry["sha256"]].append(entry["path"])
    write(output / "inventory.json", inventory)
    write(output / "selection.json", selected)
    write(output / "sample_duplicates.json", [v for v in hashes.values() if len(v) > 1])
    for name in ("SOURCES.md", "manifest.json"):
        source = root.parent / name
        if source.exists():
            (output / ("source_" + name)).write_bytes(source.read_bytes())
    summary = {
        "parameters": {
            k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()
        },
        "files": len(inventory),
        "bytes": sum(e["bytes"] for e in inventory),
        "eligible": sum(e["eligible"] for e in inventory),
        "selected": len(selected),
        "datasets": {},
        "imports": {},
        "evaluation_tasks": {},
        "calculations": {},
        "comparisons": {},
        "harness_errors": {},
        "limitations": [
            "Stratified sample, not exhaustive",
            "Original source files are only read",
            "No reference files supplied; external references fail closed",
            "Saved caches may be stale and engine semantics can differ",
            "No AI or screenshot validation; supplementary to manual acceptance",
            "SHA256 duplicate check covers selected files only",
        ],
    }
    for dataset in sorted({e["dataset"] for e in inventory}):
        entries = [e for e in inventory if e["dataset"] == dataset]
        summary["datasets"][dataset] = {
            "files": len(entries),
            "bytes": sum(e["bytes"] for e in entries),
            "extensions": dict(Counter(e["extension"] for e in entries)),
            "eligible": sum(e["eligible"] for e in entries),
            "selected": sum(e["dataset"] == dataset for e in selected),
        }
    write(output / "summary.json", summary)
    print(
        json.dumps(
            {
                "inventory": summary["files"],
                "eligible": summary["eligible"],
                "selected": len(selected),
            }
        ),
        flush=True,
    )
    if args.inventory_only:
        return
    store = TaskStore(
        WebConfig(
            data_dir=output / "tasks",
            max_workers=1,
            total_memory_mb=args.memory_mb,
            max_storage_mb=16384,
        )
    )
    budget = {"seconds": args.seconds, "memoryMb": args.memory_mb}
    evaluation_budget = {"seconds": args.evaluation_seconds, "memoryMb": args.memory_mb}
    owner = "corpus-review"
    try:
        for index, entry in enumerate(selected):
            report = {"file": entry, "evaluations": []}
            started = time.monotonic()
            stage = "source_read"
            try:
                data = (root / entry["path"]).read_bytes()
                stage = "import_submission"
                created = store.create(
                    owner,
                    {
                        "workbook": {
                            "name": "corpus" + entry["extension"],
                            "data": base64.b64encode(data).decode(),
                        },
                        "budget": budget,
                    },
                )
                stage = "import_wait"
                imported = wait(store, created["task"], owner)
                report["import"] = {k: v for k, v in imported.items() if k != "result"}
                status = imported["status"]
                summary["imports"][status] = summary["imports"].get(status, 0) + 1
                if status == "succeeded":
                    stage = "graph_snapshot"
                    graph = store.snapshot(owner, created["project"]["id"])["graph"]
                    report["graph"] = {
                        "nodes": len(graph["nodes"]),
                        "edges": len(graph["edges"]),
                        "sheets": graph["meta"]["sheets"],
                        "kinds": dict(Counter(n["kind"] for n in graph["nodes"])),
                    }
                    for node in sample_nodes(graph["nodes"], args.nodes):
                        stage = "evaluation"
                        task = wait(
                            store,
                            store.submit(
                                owner,
                                created["project"]["id"],
                                "evaluate",
                                node["id"],
                                evaluation_budget,
                            ),
                            owner,
                        )
                        result = task.get("result") or {}
                        stage = "cache_comparison"
                        verdict = comparison(node, result)
                        report["evaluations"].append(
                            {"node": node, "task": task, "comparison_review": verdict}
                        )
                        for key, value in (
                            ("evaluation_tasks", task["status"]),
                            ("calculations", result.get("status", "task_failed")),
                            ("comparisons", verdict),
                        ):
                            summary[key][value] = summary[key].get(value, 0) + 1
            except Exception as exc:
                record_exception(summary, report, stage, exc)
            report["elapsed_seconds"] = time.monotonic() - started
            write(output / f"file_{index:04d}.json", report)
            write(output / "summary.json", summary)
            print(
                json.dumps(
                    {
                        "index": index + 1,
                        "total": len(selected),
                        "file": entry["path"],
                        "seconds": round(report["elapsed_seconds"], 2),
                        "import": report.get("import", {}).get("status"),
                        "evaluations": len(report["evaluations"]),
                    }
                ),
                flush=True,
            )
    finally:
        store.close()
    summary["source_integrity"] = {
        "checked": len(selected),
        "changed": [
            entry["path"]
            for entry in selected
            if hashlib.sha256((root / entry["path"]).read_bytes()).hexdigest()
            != entry["sha256"]
        ],
    }
    summary["complete"] = True
    write(output / "summary.json", summary)


if __name__ == "__main__":
    main()
