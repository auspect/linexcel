"""Measure local-path large-workbook imports and bounded graph views."""

from __future__ import annotations

import argparse
import http.cookiejar
import json
import subprocess
import time
import urllib.error
import urllib.request
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlencode, urlsplit


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument(
        "--project",
        help="Reuse an existing immutable index; retain the earlier import report separately.",
    )
    parser.add_argument(
        "--data-dir", type=Path, default=Path("validation_screenshots/large-server")
    )
    parser.add_argument("--seconds", type=int, default=1800)
    args = parser.parse_args()
    root = Path("validation_screenshots/large") / datetime.now(UTC).strftime(
        "%Y%m%dT%H%M%S%fZ"
    )
    root.mkdir(parents=True)
    report: dict[str, Any] = {
        "workbook": str(args.workbook),
        "fileBytes": args.workbook.stat().st_size,
        "status": "running",
        "requests": [],
    }
    path = root / "validation.json"

    def save():
        path.write_text(
            json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    raw = (
        subprocess.check_output(
            [
                "uv",
                "run",
                "--extra",
                "screenshots",
                "linexcel",
                "session-link",
                "--data-dir",
                str(args.data_dir),
            ]
        )
        .decode()
        .strip()
    )
    url = urlsplit(raw)
    base = f"{url.scheme}://{url.netloc}/"
    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar())
    )

    def request(endpoint, body=None):
        start = time.perf_counter()
        data = None if body is None else json.dumps(body).encode()
        response = opener.open(
            urllib.request.Request(
                base + "api/" + endpoint,
                data=data,
                headers={"Content-Type": "application/json"},
            ),
            timeout=60,
        )
        payload = response.read()
        elapsed = time.perf_counter() - start
        if not endpoint.startswith("tasks/"):
            report["requests"].append(
                {"endpoint": endpoint, "seconds": elapsed, "bytes": len(payload)}
            )
        return json.loads(payload)

    def wait(task):
        deadline = time.monotonic() + args.seconds + 30
        while task["status"] in {"queued", "running"} and time.monotonic() < deadline:
            time.sleep(1)
            task = request("tasks/" + task["id"])["task"]
        save()
        return task

    request("session", {"token": url.fragment.split("=", 1)[1]})
    budget = {"seconds": args.seconds, "memoryMb": 2048}
    start = time.perf_counter()
    created = (
        {"project": {"id": args.project}}
        if args.project
        else request(
            "projects/local",
            {
                "path": str(args.workbook.resolve()),
                "name": args.workbook.stem,
                "budget": budget,
            },
        )
    )
    project = created["project"]["id"]
    report["project"] = project
    save()
    imported = (
        {"status": "reused_immutable_index"} if args.project else wait(created["task"])
    )
    report["import"] = {"seconds": time.perf_counter() - start, "task": imported}
    save()
    if imported["status"] not in {"succeeded", "reused_immutable_index"}:
        report["status"] = "failed_import"
        save()
        print(path)
        return 1
    start = time.perf_counter()
    built = wait(
        request(
            f"projects/{project}/tasks", {"operation": "build_graph", "budget": budget}
        )["task"]
    )
    report["buildGraph"] = {"seconds": time.perf_counter() - start, "task": built}
    save()
    if built["status"] != "succeeded":
        report["status"] = "failed_graph"
        save()
        print(path)
        return 1
    views = []
    endpoint = f"projects/{project}/graph-view"
    view = request(endpoint)
    for _ in range(5):
        request(endpoint)
    for _ in range(12):
        views.append(view)
        child = next(
            (
                n
                for n in view["nodes"]
                if n.get("viewId") is not None and n["kind"] != "outside"
            ),
            None,
        )
        if child is None:
            break
        view = request(endpoint + "?" + urlencode({"view": child["viewId"]}))
    report["views"] = views
    formula = request(f"projects/{project}/nodes?kind=cell&limit=1")["nodes"]
    if formula:
        report["selectedNode"] = formula[0]["id"]
        start = time.perf_counter()
        result = wait(
            request(
                f"projects/{project}/tasks",
                {
                    "operation": "evaluate",
                    "nodeId": formula[0]["id"],
                    "force": False,
                    "budget": budget,
                },
            )["task"]
        )
        report["calculation"] = {"seconds": time.perf_counter() - start, "task": result}
    report["status"] = (
        "passed" if all(len(v["nodes"]) <= 33 for v in views) else "failed_payload_cap"
    )
    save()
    print(path)
    return int(report["status"] != "passed")


if __name__ == "__main__":
    raise SystemExit(main())
