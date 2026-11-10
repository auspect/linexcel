"""What a question's tool results look like to the model.

The full result remains the reader's evidence: the answer is reread from the
source indexes and keeps every field. A model only chooses the next call, so it
receives identities, labels, saved values, formulas and stop reasons, bounded
per result. Repeating a 64-character revision hash and the XML provenance on
every cell made one lineage cost more tokens than a whole question.
"""

from __future__ import annotations

import json
from urllib.parse import unquote

TEXT_CHARS = 300
FORMULA_CHARS = 600
LABEL_CHARS = 120
LABELS = 4
OCCURRENCES = 5
CACHED_VALUES = 8
# One delivered result. Sixteen of them stay well inside the run's byte budget.
RESULT_BYTES = 24 * 1024
LIST_ITEMS = {
    "results": 20,
    "cells": 25,
    "links": 20,
    "nodes": 60,
    "edges": 120,
    "boundaries": 40,
}


def _size(value) -> int:
    return len(
        json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    )


def _clip(value, limit=TEXT_CHARS):
    if isinstance(value, str) and len(value) > limit:
        return value[:limit] + "…"
    return value


def _labels(labels):
    result = []
    for label in (labels or [])[:LABELS]:
        if not isinstance(label, dict) or not label.get("text"):
            continue
        item = {"text": _clip(label["text"], LABEL_CHARS)}
        if label.get("role"):
            item["role"] = label["role"]
        cell = label.get("sourceCell") or label.get("cell")
        if cell:
            item["cell"] = cell
        result.append(item)
    return result


def _cell(raw, many):
    cell = {}
    if raw.get("handle"):
        cell["handle"] = raw["handle"]
    if many and raw.get("filename"):
        cell["file"] = raw["filename"]
    cell.update(
        sheet=raw.get("sheet"),
        address=raw.get("address"),
        kind=raw.get("kind"),
        value=_clip(raw.get("value")),
    )
    if raw.get("formula"):
        cell["formula"] = _clip(raw["formula"], FORMULA_CHARS)
    labels = _labels(raw.get("labels"))
    if labels:
        cell["labels"] = labels
    for key in ("valueProvenance", "cacheStatus"):
        if raw.get(key):
            cell[key] = raw[key]
    return cell


def _hit(raw, many):
    hit = _cell(raw, many)
    if raw.get("name"):
        hit["name"] = _clip(raw["name"], LABEL_CHARS)
    count = raw.get("occurrenceCount") or 0
    if count > 1:
        occurrences = [
            {"address": item.get("address"), "handle": item["handle"]}
            for item in raw.get("occurrences") or []
            if item.get("handle")
        ]
        hit["occurrenceCount"] = count
        hit["occurrences"] = occurrences[:OCCURRENCES]
    return hit


def _coverage(coverage, keys):
    coverage = coverage or {}
    return {key: coverage[key] for key in keys if key in coverage}


def _link(fact):
    link = {
        key: _clip(fact[key], LABEL_CHARS)
        for key in ("filename", "reference", "sheet", "address", "status", "sourceNode")
        if fact.get(key)
    }
    source, target = fact.get("sourceRevision"), fact.get("targetRevision")
    if isinstance(source, dict):
        link["sourceProject"] = source.get("projectId")
    if isinstance(target, dict):
        link["targetProject"] = target.get("projectId")
    link["candidates"] = len(fact.get("candidates") or [])
    link["pinned"] = bool(fact.get("pin"))
    if fact.get("supported") is False:
        link["supported"] = False
    if fact.get("versionCaveat"):
        link["versionCaveat"] = _clip(str(fact["versionCaveat"]), LABEL_CHARS)
    cached = _cached(fact.get("cached"))
    if cached:
        # Saved by Excel in the citing workbook; the linked file was not read.
        link["cachedInCitingWorkbook"] = cached
    return link


def _cached(items):
    return [
        {
            "row": item.get("row"),
            "column": item.get("column"),
            "value": _clip(item.get("value")),
        }
        for item in (items or [])[:CACHED_VALUES]
        if isinstance(item, dict)
    ]


def _boundary(raw):
    boundary = {"node": raw.get("node"), "reason": raw.get("reason")}
    path = raw.get("relationshipPath")
    if isinstance(path, str) and path:
        # The file name only: its folder is on the workbook author's machine.
        name = unquote(path.replace("\\", "/").rsplit("/", 1)[-1])
        boundary["filename"] = _clip(name, LABEL_CHARS)
    for key, value in raw.items():
        if key in boundary or key in {"component", "relationshipPath", "linkId"}:
            continue
        if isinstance(value, (str, int, float, bool)):
            boundary[key] = _clip(value, LABEL_CHARS)
    if isinstance(raw.get("component"), list):
        boundary["componentNodes"] = len(raw["component"])
    cached = _cached(raw.get("cached"))
    if cached:
        boundary["cachedInCitingWorkbook"] = cached
    return boundary


def _project(name, data, many):
    if name in {"list_workbooks", "list_revisions"}:
        return {
            "workbooks": [
                {
                    "project": item.get("projectId"),
                    "filename": item.get("filename"),
                    "sheets": item.get("sheets", []),
                }
                for item in data.get("revisions", [])
            ]
        }
    if name == "search_catalog":
        pagination = data.get("pagination") or {}
        view = {
            "results": [_hit(item, many) for item in data.get("results", [])],
            "matched": pagination.get("matched"),
            "hasMore": pagination.get("hasMore"),
            "coverage": _coverage(
                data.get("coverage"),
                (
                    "complete",
                    "capsHit",
                    "nodesTotal",
                    "nodesScanned",
                    "nodesUnscanned",
                    "numericPolicy",
                ),
            ),
        }
        if data.get("status"):
            view["status"] = data["status"]
        return view
    if name == "describe_sheet":
        return {
            "sheet": data.get("sheet"),
            "sampleOnly": True,
            "cells": [_cell(item, many) for item in data.get("cells", [])],
        }
    if name == "read_range":
        return {
            "cells": [_cell(item, many) for item in data.get("cells", [])],
            "blankOrUnindexed": data.get("blankOrUnindexed"),
            "sourceCompleteness": data.get("sourceCompleteness"),
        }
    if name == "get_cell":
        return _cell(data, many)
    if name == "find_value":
        return {
            "cells": [_cell(item, many) for item in data.get("cells", [])],
            "coverage": data.get("coverage"),
        }
    if name == "trace_lineage":
        return {
            "direction": data.get("direction"),
            "nodes": [_cell(item, many) for item in data.get("nodes", [])],
            "edges": [
                {
                    "source": item.get("source"),
                    "target": item.get("target"),
                    "kind": item.get("kind"),
                }
                for item in data.get("edges", [])
            ],
            # A constant input ends a branch by nature; its node already says so.
            "boundaries": [
                _boundary(item)
                for item in data.get("boundaries", [])
                if item.get("reason") != "input_constant"
            ],
            "coverage": _coverage(
                data.get("coverage"),
                (
                    "complete",
                    "exhaustive",
                    "capsHit",
                    "boundaries",
                    "nodesReturned",
                    "edgesReturned",
                    "depth",
                ),
            ),
        }
    if name == "list_external_links":
        pagination = data.get("pagination") or {}
        return {
            "links": [_link(item) for item in data.get("links", [])],
            "matched": pagination.get("matched"),
            "hasMore": pagination.get("hasMore"),
            "coverage": data.get("coverage"),
        }
    if name == "compute":
        view = {
            key: data.get(key)
            for key in (
                "computeHandle",
                "op",
                "value",
                "error",
                "blanks",
                "missingSavedValues",
            )
        }
        # The run reduces the receipt to its identity; the recipe stays local.
        view["receipt"] = data.get("receipt")
        view["operands"] = [
            {
                "handle": item.get("handle"),
                "sheet": item.get("sheet"),
                "address": item.get("address"),
                "value": _clip(item.get("value")),
            }
            for item in data.get("operands", [])
        ]
        return view
    return data


def _keep_lineage_consistent(view):
    """Edges and boundaries may only name nodes the model was given."""
    kept = {node.get("handle") for node in view["nodes"]}
    view["edges"] = [
        edge
        for edge in view["edges"]
        if edge["source"] in kept and edge["target"] in kept
    ]
    view["boundaries"] = [item for item in view["boundaries"] if item["node"] in kept]


def model_view(name, data, *, many=False, budget=RESULT_BYTES):
    """Return ``(view, omitted)``: the compact result and what was left out.

    ``omitted`` maps a list name to the number of entries the model did not
    receive. A non-empty value means the model saw part of the result, which the
    run records as an incomplete search.
    """
    view = _project(name, data, many)
    if not isinstance(view, dict):
        return view, {}
    before = {
        key: len(view[key]) for key in LIST_ITEMS if isinstance(view.get(key), list)
    }
    lineage = name == "trace_lineage"
    # Entries are dropped from the end: results are ordered, and a lineage lists
    # its nodes breadth first, so what remains is the part nearest the request.
    order = ("nodes", "results", "cells", "links", "edges", "boundaries")
    for key in order:
        if key in before:
            del view[key][LIST_ITEMS[key] :]
            if lineage and key == "nodes":
                _keep_lineage_consistent(view)
    size = _size(view)
    while size > budget:
        key = next((k for k in order if isinstance(view.get(k), list) and view[k]), "")
        if not key:
            break
        view[key].pop()
        if lineage and key == "nodes":
            _keep_lineage_consistent(view)
        size = _size(view)
    omitted = {
        key: count - len(view[key])
        for key, count in before.items()
        if count > len(view[key])
    }
    if omitted:
        view["omittedForModel"] = omitted
        view["hint"] = (
            "Part of this result was left out to keep the question within its "
            "budget. Narrow the call: a smaller depth, a more specific query or "
            "a smaller range."
        )
    return view, omitted
