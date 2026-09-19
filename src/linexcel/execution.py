"""Process isolation for workbook analysis; no native engine crosses the boundary."""

from __future__ import annotations

import json
import math
import os
import signal
import subprocess
import sys
import tempfile
import time
import uuid
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, cast


@dataclass(frozen=True)
class ExecutionPolicy:
    """Analysis budgets, excluding source reading, AI, screenshots and export.

    ``isolated=False`` explicitly opts into the legacy live-engine API, without
    a hard time or memory guarantee. The default never retries in that mode.
    """

    seconds: float = 120.0
    memory_mb: int = 2048
    isolated: bool = True

    def __post_init__(self):
        if isinstance(self.seconds, bool) or not isinstance(self.seconds, (int, float)):
            raise TypeError("seconds must be a finite non-negative number")
        if not math.isfinite(self.seconds) or self.seconds < 0:
            raise ValueError("seconds must be a finite non-negative number")
        if isinstance(self.memory_mb, bool) or not isinstance(self.memory_mb, int):
            raise TypeError("memory_mb must be a positive integer")
        if self.memory_mb <= 0:
            raise ValueError("memory_mb must be a positive integer")
        if not isinstance(self.isolated, bool):
            raise TypeError("isolated must be boolean")


def add_coverage(graph: dict[str, Any]) -> None:
    """Count graph nodes, never extrapolate representative cells to a workbook."""
    categories: dict[str, Any] = {
        key: {"count": 0, "nodeIds": []}
        for key in ("engine", "cache", "unavailable", "other", "divergent")
    }
    for node in graph.get("nodes", []):
        source = node.get("valueSource")
        formula_node = node.get("kind") in {"cell", "group"}
        key = "other"
        if formula_node and node.get("value") is None:
            key = "unavailable"
        elif formula_node and source == "engine":
            key = "engine"
        elif (
            source in {"file", "volatile", "external-cache"}
            and node.get("value") is not None
        ):
            key = "cache"
        categories[key]["nodeIds"].append(node["id"])
        if (
            node.get("cachedAgreement") == "differ"
            or node.get("groupCachedAgreement") == "differ"
        ):
            categories["divergent"]["nodeIds"].append(node["id"])
    for value in categories.values():
        value["count"] = len(value["nodeIds"])
    meta = graph.setdefault("meta", {})
    meta["coverage"] = {
        "scope": "graph_nodes",
        "totalNodes": len(graph.get("nodes", [])),
        "categories": categories,
        "omissions": {
            "status": "reported"
            if meta.get("analysisCoverage", {}).get("omissions")
            else "not_certified",
            "items": meta.get("analysisCoverage", {}).get("omissions", []),
            "warnings": meta.get("warnings", []),
        },
    }


def _read_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _checkpoint(root: Path) -> dict:
    paths = sorted(root.glob("checkpoint-*.json"))
    return _read_json(paths[-1]) if paths else {}


@contextmanager
def _worker_directory():
    directory = tempfile.TemporaryDirectory(prefix="linexcel-")
    try:
        yield directory.name
    finally:
        # Job termination can precede release of inherited file handles by a
        # few milliseconds. Only retry these Windows sharing/locking errors;
        # cleanup must really succeed, or the original error reaches the caller.
        deadline = time.monotonic() + 2
        while True:
            try:
                directory.cleanup()
                break
            except OSError as exc:
                if (
                    os.name != "nt"
                    or getattr(exc, "winerror", None) not in {32, 33}
                    or time.monotonic() >= deadline
                ):
                    raise
                time.sleep(0.005)


def run_isolated(data: bytes, kwargs: dict, policy: ExecutionPolicy) -> dict:
    """Run one worker, enforcing elapsed time even inside a native call."""
    started = time.monotonic()
    status = "timed_out"
    checkpoint: dict = {}
    payload: dict = {}
    exit_code = None
    diagnostic = ""
    with _worker_directory() as directory:
        root = Path(directory)
        (root / "input.xlsx").write_bytes(data)
        (root / "request.json").write_text(
            json.dumps(
                {
                    "kwargs": kwargs,
                    "memory_mb": policy.memory_mb,
                },
                default=str,
            ),
            encoding="utf-8",
        )
        if policy.seconds > time.monotonic() - started:
            with (root / "stderr.txt").open("w", encoding="utf-8") as log:
                command = [sys.executable, "-m", "linexcel._worker", str(root)]
                if os.name != "nt":
                    # Apply the virtual-address limit before importing linexcel.
                    bootstrap = (
                        "import resource,runpy,sys; "
                        "resource.setrlimit(resource.RLIMIT_AS,"
                        f"({policy.memory_mb * 1048576},)*2); "
                        "sys.argv=['linexcel._worker',sys.argv[1]]; "
                        "runpy.run_module('linexcel._worker',run_name='__main__')"
                    )
                    command = [sys.executable, "-c", bootstrap, str(root)]
                process = subprocess.Popen(
                    command,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=log,
                    start_new_session=os.name != "nt",
                    # Suspend before Python imports anything; attach the Job first.
                    creationflags=(subprocess.CREATE_NO_WINDOW | 0x4)
                    if os.name == "nt"
                    else 0,
                )
                job = None
                try:
                    if os.name == "nt":
                        from linexcel._windows_job import WorkerJob

                        job = WorkerJob(process, policy.memory_mb)
                        job.resume(process)
                    (root / "ready").touch()
                    last_phase = None
                    while process.poll() is None:
                        checkpoint = _checkpoint(root)
                        phase = checkpoint.get("phase", "starting worker")
                        if kwargs.get("verbose") and phase != last_phase:
                            elapsed = time.monotonic() - started
                            print(
                                f"[linexcel] {phase}; {elapsed:.1f}s elapsed; "
                                f"{max(0, policy.seconds - elapsed):.1f}s "
                                "budget remaining",
                                file=sys.stderr,
                            )
                            last_phase = phase
                        if time.monotonic() - started >= policy.seconds:
                            break
                        time.sleep(
                            min(
                                0.05,
                                max(0, policy.seconds - (time.monotonic() - started)),
                            )
                        )
                    else:
                        payload = _read_json(root / "result.json")
                        status = payload.get("status", "crashed")
                except KeyboardInterrupt:
                    status = "cancelled"
                finally:
                    if job is not None:
                        job.close()  # Kill descendants even after worker exit.
                    elif os.name != "nt":
                        try:
                            os.killpg(process.pid, signal.SIGKILL)
                        except ProcessLookupError:
                            pass
                    elif process.poll() is None:
                        process.kill()
                    process.wait(timeout=5)
                checkpoint = _checkpoint(root)
                if kwargs.get("verbose"):
                    for phase, metric in checkpoint.get("phaseMetrics", {}).items():
                        print(
                            f"[linexcel] {phase}: completed in "
                            f"{metric['elapsedSeconds']:.3f}s",
                            file=sys.stderr,
                        )
                    print(
                        f"[linexcel] total: {time.monotonic() - started:.3f}s; "
                        f"{status}",
                        file=sys.stderr,
                    )
                exit_code = process.returncode
                if status != "completed":
                    with (root / "stderr.txt").open("rb") as errors:
                        errors.seek(max(0, (root / "stderr.txt").stat().st_size - 4096))
                        diagnostic = errors.read().decode("utf-8", errors="replace")
                    if (
                        status == "crashed"
                        and exit_code is not None
                        and (exit_code & 0xFFFFFFFF) in {0xC0000017, 0xC000012D}
                    ):
                        status = "memory_limit"
        if status == "error":
            raise ValueError(payload.get("error", "Analysis worker failed"))
        graph: Any = payload.get("graph") if status == "completed" else None
        if graph is None:
            graph = cast(
                dict[str, Any],
                checkpoint.get("graph")
                or {
                    "meta": {
                        "filename": kwargs.get("filename", "workbook.xlsx"),
                        "stats": {
                            "totalNodes": 0,
                            "totalEdges": 0,
                            "totalFormulas": None,
                        },
                        "warnings": [],
                    },
                    "nodes": [],
                    "edges": [],
                    "sheets": checkpoint.get("sheets", []),
                },
            )
            graph["meta"].setdefault("warnings", []).append(
                f"Analysis {status} during {checkpoint.get('phase', 'startup')}. "
                "The requested analysis is incomplete; "
                "no unfinished engine values are published."
            )
            graph["meta"]["sourceEvidence"] = checkpoint.get("sourceEvidence", {})
            graph["meta"]["analysisCoverage"] = {
                "requested": "targeted_static_closure"
                if kwargs.get("targets")
                else "workbook",
                "extractedFormulaCells": None,
                "dependencyCompleteness": "not_inspected",
                "omissions": [
                    {"phase": checkpoint.get("phase", "startup"), "reason": status}
                ],
            }
        graph["meta"]["execution"] = {
            "status": status,
            "isolated": True,
            "budgetSeconds": policy.seconds,
            "elapsedSeconds": round(time.monotonic() - started, 3),
            "memoryLimitMiB": policy.memory_mb,
            "memoryLimitKind": "job_commit"
            if os.name == "nt"
            else "per_process_virtual_address",
            "phase": checkpoint.get("phase", "startup"),
            "completedPhases": checkpoint.get("completedPhases", []),
            "phaseMetrics": checkpoint.get("phaseMetrics", {}),
            "engineAvailable": False,
            "exitCode": exit_code,
            "diagnostic": diagnostic,
        }
        add_coverage(graph)
        return {"graph": graph, "engine": None, "analysisId": uuid.uuid4().hex[:16]}
