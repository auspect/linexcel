"""Throwaway snapshot tool for the analyzer.py refactor.

Generates a canonical JSON snapshot of analyze_workbook()'s graph output for a
small set of fixtures, with non-deterministic fields (analysisId, analyzedAt)
stripped and everything else sorted so the diff is stable across runs.

Usage:
    uv run python scripts/freeze_graph.py /tmp/snapshot_before.json
    uv run python scripts/freeze_graph.py /tmp/snapshot_after.json
    diff /tmp/snapshot_before.json /tmp/snapshot_after.json
"""

import io
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tests"))

from linexcel.analyzer import analyze_workbook  # noqa: E402

FIXTURES_DIR = Path(__file__).resolve().parent.parent / "tests" / "fixtures"


def _lineage_workbook() -> bytes:
    from conftest import build_lineage_workbook

    return build_lineage_workbook()


def _sources() -> dict[str, bytes]:
    sources = {"lineage": _lineage_workbook()}
    for name in ("macros.xlsm", "power_query.xlsx"):
        path = FIXTURES_DIR / name
        if path.exists():
            sources[name] = path.read_bytes()
    return sources


def _strip_nondeterministic(graph: dict) -> dict:
    meta = graph.get("meta", {})
    meta.pop("analyzedAt", None)
    return graph


def _canonical(obj):
    """Sort dicts by key and sort lists of dicts by their JSON text, so
    ordering that is an implementation detail (dict iteration, node insertion
    order) doesn't show up as a diff."""
    if isinstance(obj, dict):
        return {k: _canonical(obj[k]) for k in sorted(obj.keys(), key=str)}
    if isinstance(obj, list):
        items = [_canonical(v) for v in obj]
        try:
            items.sort(key=lambda v: json.dumps(v, sort_keys=True, default=str))
        except TypeError:
            pass
        return items
    return obj


def main() -> None:
    if len(sys.argv) != 2:
        print("usage: freeze_graph.py <output.json>", file=sys.stderr)
        raise SystemExit(2)
    out_path = Path(sys.argv[1])

    snapshot: dict[str, object] = {}
    for name, data in _sources().items():
        result = analyze_workbook(data, filename=name)
        graph = _strip_nondeterministic(result["graph"])
        snapshot[name] = _canonical(graph)

    out_path.write_text(
        json.dumps(snapshot, indent=2, sort_keys=True, default=str), encoding="utf-8"
    )
    print(f"wrote {out_path}", file=sys.stderr)


if __name__ == "__main__":
    main()
