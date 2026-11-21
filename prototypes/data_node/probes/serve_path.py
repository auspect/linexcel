"""Drive one workbook through the real `linexcel serve` coordinator.

Same functions as the UI: local import (uploads.local), import task, the
default graph view (build_graph task + graph-view), node/neighborhood reads,
the Data tab (Datasette snapshot + Excel presentation) and one evaluation.
Default WebConfig limits unless overridden; UI default budgets (600 s, 1024 MiB).

usage: serve_path.py <case-name> <xlsx> <data-dir> <out.json>
         [refs=dir1;dir2] [cell=Sheet!A1] [seconds=600] [memory=1024]
         [storage_mb=2048] [upload_mb=256] [skip=graph,data]
"""

import hashlib
import io
import json
import sys
import time
import urllib.parse
from pathlib import Path

from linexcel.web import create_app

name, workbook, data_dir, out = sys.argv[1:5]
opts = dict(a.split("=", 1) for a in sys.argv[5:])
seconds = int(opts.get("seconds", 600))
memory = int(opts.get("memory", 1024))
skip = set(filter(None, opts.get("skip", "").split(",")))
config = {}
if "storage_mb" in opts:
    config["max_storage_mb"] = int(opts["storage_mb"])
if "upload_mb" in opts:
    config["max_upload_mb"] = int(opts["upload_mb"])
app = create_app(Path(data_dir), resolve_user=lambda env: env.get("user"), **config)
owner = hashlib.sha256(b"alice").hexdigest()
report = {
    "case": name,
    "workbook": Path(workbook).name,
    "config": config,
    "budget": {"seconds": seconds, "memoryMb": memory},
}


def call(method, path, body=None, query=""):
    data = json.dumps(body or {}).encode()
    status = []
    started = time.perf_counter()
    raw = b"".join(
        app(
            {
                "REQUEST_METHOD": method,
                "PATH_INFO": path,
                "QUERY_STRING": query,
                "CONTENT_TYPE": "application/json",
                "CONTENT_LENGTH": str(len(data)),
                "wsgi.input": io.BytesIO(data),
                "user": "alice",
            },
            lambda s, h: status.append(s),
        )
    )
    elapsed = round(time.perf_counter() - started, 3)
    try:
        payload = json.loads(raw)
    except ValueError:
        payload = {"_raw": raw[:300].decode("utf-8", "replace")}
    return int(status[0].split()[0]), payload, elapsed, len(raw)


def wait(task_id, limit):
    deadline = time.monotonic() + limit
    while time.monotonic() < deadline:
        task = app.store.get_task(owner, task_id)
        if task["status"] in {"failed", "succeeded", "cancelled"}:
            return task
        time.sleep(0.25)
    return {"status": "harness_timeout"}


def summary(task):
    keep = {k: task.get(k) for k in ("status", "error", "operation")}
    result = task.get("result") or {}
    if isinstance(result, dict):
        keep["result"] = {
            k: result.get(k) for k in ("nodeCount", "indexStatus") if k in result
        }
        diags = result.get("diagnostics") or []
        keep["diagnosticCount"] = len(diags)
        keep["diagnosticSample"] = diags[:6]
    return keep


try:
    started = time.perf_counter()
    try:
        refs = [p for p in opts.get("refs", "").split(";") if p]
        created = app.store.uploads.local(
            owner,
            {
                "path": str(Path(workbook).resolve()),
                "references": refs,
                "budget": {"seconds": seconds, "memoryMb": memory},
            },
        )
    except Exception as exc:
        report["upload"] = {
            "status": "refused",
            "error": f"{type(exc).__name__}: {getattr(exc, 'status', '')} {exc}",
        }
        raise SystemExit
    project = created["project"]["id"]
    task = wait(created["task"]["id"], seconds + 120)
    report["importSeconds"] = round(time.perf_counter() - started, 1)
    report["import"] = summary(task)
    code, snap, t, size = call("GET", f"/api/projects/{project}")
    meta = (snap.get("graph") or {}).get("meta") or snap.get("meta") or {}
    report["snapshot"] = {
        "code": code,
        "seconds": t,
        "bytes": size,
        "error": snap.get("error"),
    }
    for key in ("status", "nodeCount", "formulaCount", "inputCount"):
        if key in meta:
            report["snapshot"][key] = meta[key]
    details = meta.get("sheetDetails") or []
    report["sheets"] = [
        {
            "name": s.get("name"),
            "status": s.get("status"),
            "diagnostics": (s.get("diagnostics") or [])[:2],
        }
        for s in details
    ]
    report["indexDiagnostics"] = (meta.get("diagnostics") or [])[:8]
    sheets = meta.get("sheets") or [s["name"] for s in details]
    code, nodes, t, size = call(
        "GET", f"/api/projects/{project}/nodes", query="limit=50"
    )
    report["nodesPage"] = {
        "code": code,
        "seconds": t,
        "count": len(nodes.get("nodes", [])) if isinstance(nodes, dict) else None,
        "error": nodes.get("error") if isinstance(nodes, dict) else None,
    }
    cell = opts.get("cell")
    if cell:
        q = urllib.parse.urlencode({"nodeId": cell, "depth": 2, "limit": 200})
        code, around, t, size = call(
            "GET", f"/api/projects/{project}/neighborhood", query=q
        )
        report["neighborhood"] = {
            "code": code,
            "seconds": t,
            "nodes": len(around.get("nodes", [])) if code == 200 else None,
            "truncated": (around.get("pagination") or {}).get("truncated")
            if code == 200
            else None,
            "error": around.get("error"),
        }
    if "graph" not in skip and report["import"]["status"] == "succeeded":
        started = time.perf_counter()
        code, submitted, t, _ = call(
            "POST",
            f"/api/projects/{project}/tasks",
            {
                "operation": "build_graph",
                "budget": {"seconds": seconds, "memoryMb": memory},
            },
        )
        if code == 200:
            graph_task = wait(submitted["task"]["id"], seconds + 120)
            report["buildGraph"] = summary(graph_task) | {
                "seconds": round(time.perf_counter() - started, 1)
            }
            code, view, t, size = call(
                "GET", f"/api/projects/{project}/graph-view", query=""
            )
            report["graphView"] = {
                "code": code,
                "seconds": t,
                "bytes": size,
                "error": view.get("error") if isinstance(view, dict) else None,
            }
        else:
            report["buildGraph"] = {"code": code, "error": submitted.get("error")}
    if "data" not in skip and report["import"]["status"] == "succeeded":
        code, data, t, size = call("GET", f"/api/projects/{project}/data")
        report["dataTab"] = {
            "code": code,
            "seconds": t,
            "error": data.get("error"),
            "sheets": [
                (s.get("sheet"), s.get("layout"), s.get("cellCount"))
                for s in data.get("sheets", [])
            ][:12]
            if code == 200
            else None,
        }
        storage = sum(
            p.stat().st_size for p in Path(data_dir).rglob("*") if p.is_file()
        )
        report["storageAfterDataMB"] = round(storage / 1e6, 1)
        presentation = []
        for sheet in sheets[:40]:
            q = urllib.parse.urlencode(
                {"sheet": sheet, "row": 1, "column": 1, "rows": 60, "columns": 20}
            )
            code, view, t, size = call(
                "GET", f"/api/projects/{project}/data/presentation", query=q
            )
            presentation.append(
                {
                    "sheet": sheet,
                    "code": code,
                    "seconds": t,
                    "cells": len(view.get("cells", [])) if code == 200 else None,
                    "error": view.get("error"),
                }
            )
        report["presentation"] = presentation
    if cell and report["import"]["status"] == "succeeded":
        started = time.perf_counter()
        code, submitted, t, _ = call(
            "POST",
            f"/api/projects/{project}/tasks",
            {
                "operation": "evaluate",
                "nodeId": cell,
                "budget": {"seconds": seconds, "memoryMb": memory},
            },
        )
        if code == 200:
            done = wait(submitted["task"]["id"], seconds + 120)
            code, result, t, _ = call(
                "GET", f"/api/tasks/{submitted['task']['id']}/export"
            )
            r = result.get("result", result) if isinstance(result, dict) else {}
            report["evaluate"] = {
                "task": done.get("status"),
                "error": done.get("error"),
                "seconds": round(time.perf_counter() - started, 1),
                "status": r.get("status"),
                "comparison": r.get("comparison"),
                "value": r.get("value"),
                "cached": r.get("cachedValue"),
                "diagnostics": (r.get("diagnostics") or [])[:4],
            }
        else:
            report["evaluate"] = {"code": code, "error": submitted.get("error")}
except SystemExit:
    pass
except BaseException as exc:
    report["harnessError"] = f"{type(exc).__name__}: {exc}"
finally:
    try:
        report["storageMB"] = round(
            sum(p.stat().st_size for p in Path(data_dir).rglob("*") if p.is_file())
            / 1e6,
            1,
        )
    except OSError:
        pass
    app.close()
    Path(out).write_text(
        json.dumps(report, indent=1, ensure_ascii=False, default=str), "utf-8"
    )
    print(json.dumps(report, ensure_ascii=False, default=str)[:3000])
