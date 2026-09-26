"""Process isolation for workbook analysis; no native engine crosses the boundary."""

from __future__ import annotations

import json
import math
import os
import platform
import re
import signal
import subprocess
import sys
import tempfile
import time
import uuid
from contextlib import contextmanager
from dataclasses import dataclass, replace
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Any, cast

DIAGNOSTIC_BYTES = 16_384


def _runtime_versions() -> dict[str, str]:
    versions = {"python": platform.python_version(), "platform": sys.platform}
    for package in ("linexcel", "formualizer"):
        try:
            versions[package] = version(package)
        except PackageNotFoundError:
            versions[package] = "unknown"
    return versions


def _failure_details(status: str, code: int | None, stderr: str) -> dict:
    """Interpret evidence, without guessing OOM from SIGKILL or every abort."""
    details: dict[str, Any] = {"kind": status, "summary": ""}
    unsigned = code & 0xFFFFFFFF if code is not None else None
    if os.name == "nt" and unsigned is not None:
        details["exitCodeHex"] = f"0x{unsigned:08X}"
    if status == "memory_limit":
        details["summary"] = "Memory allocation failed under the worker memory limit."
    if status != "crashed":
        return details
    windows = {
        0xC0000017: ("memory_limit", "Windows reported insufficient memory."),
        0xC000012D: ("memory_limit", "Windows reported the commitment limit exceeded."),
        0xC00000FD: ("stack_overflow", "Native stack overflow."),
        0xC0000005: ("access_violation", "Native access violation."),
        0xC0000409: (
            "native_abort",
            "Windows fast-fail; the root cause is not established.",
        ),
    }
    if os.name == "nt" and unsigned in windows:
        details["kind"], details["summary"] = windows[unsigned]
    elif os.name != "nt" and code is not None and code < 0:
        try:
            details["signal"] = signal.Signals(-code).name
        except ValueError:
            details["signal"] = str(-code)
        details["kind"] = "native_signal"
        details["summary"] = f"Worker terminated by {details['signal']}."
    # Rust's allocation handler aborts instead of raising Python MemoryError.
    # An abort/fast-fail code alone does not establish memory exhaustion.
    if re.search(r"(?m)^memory allocation of \d+ bytes failed\s*$", stderr):
        details.update(
            kind="memory_limit",
            summary="Native memory allocation failed under the worker memory limit.",
        )
    elif "has overflowed its stack" in stderr:
        details.update(kind="stack_overflow", summary="Native stack overflow.")
    elif "pyo3_runtime.PanicException:" in stderr or re.search(
        r"(?m)^thread .* panicked at ", stderr
    ):
        details.update(
            kind="native_panic",
            summary="Rust engine panic; inspect the diagnostic log.",
        )
    if not details["summary"]:
        details["kind"] = "unknown_exit"
        details["summary"] = (
            "Worker exited without a result; the cause is not established."
        )
    return details


@dataclass(frozen=True)
class ExecutionPolicy:
    """Analysis budgets, excluding source reading, AI, screenshots and export.

    ``isolated=False`` explicitly opts into the legacy live-engine API, without
    a hard time or memory guarantee. The default never retries in that mode.
    """

    seconds: float = 120.0
    memory_mb: int = 2048
    isolated: bool = True
    memory_retries: int = 2

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
        if isinstance(self.memory_retries, bool) or not isinstance(
            self.memory_retries, int
        ):
            raise TypeError("memory_retries must be an integer between 0 and 2")
        if not 0 <= self.memory_retries <= 2:
            raise ValueError("memory_retries must be between 0 and 2")


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
    """On proven OOM, try a smaller source inventory under the same budgets."""
    started = time.monotonic()
    result = _run_isolated_once(data, kwargs, policy)
    initial = result["graph"]["meta"]["execution"]
    if initial["status"] != "memory_limit" or not policy.memory_retries:
        return result
    if kwargs.get("targets"):
        initial["recoverySkipped"] = (
            "targeted_analysis_requires_complete_dependency_closure"
        )
        return result
    attempts = [dict(initial, mode="analysis")]
    cell_budget = min(5000, max(1, policy.memory_mb * 8))
    for retry_index in range(policy.memory_retries):
        remaining = policy.seconds - (time.monotonic() - started)
        if remaining <= 0:
            break
        recovery_kwargs = dict(kwargs)
        # Reduce effective per-sheet caps too: otherwise a 400-node cap can
        # make both a 5000-cell and a 1250-cell attempt read the same 400 cells.
        for key, default in (
            ("max_nodes_per_sheet", 400),
            ("max_cells_per_sheet", None),
        ):
            limit = kwargs.get(key)
            limit = default if limit is None else limit
            if limit is not None:
                recovery_kwargs[key] = limit // (4**retry_index)
        retry = _run_isolated_once(
            data,
            recovery_kwargs,
            replace(policy, seconds=remaining),
            recovery_cells=cell_budget,
        )
        evidence = retry["graph"]["meta"]["execution"]
        attempts.append(
            dict(
                evidence,
                mode="source_only",
                storedCellBudget=cell_budget,
                cellsPerSheet=recovery_kwargs.get("max_cells_per_sheet"),
                nodesPerSheet=recovery_kwargs.get("max_nodes_per_sheet"),
            )
        )
        if evidence["status"] == "completed":
            graph = retry["graph"]
            graph["meta"]["warnings"] = (
                result["graph"]["meta"]["warnings"] + graph["meta"]["warnings"]
            )
            graph["meta"]["execution"] = initial
            graph["meta"]["sourceEvidence"] = result["graph"]["meta"].get(
                "sourceEvidence", {}
            )
            initial["recovery"] = {
                "status": "completed",
                "mode": "source_only",
                "storedCellBudget": cell_budget,
                "recalculated": False,
            }
            result = retry
            add_coverage(graph)
            break
        if evidence["status"] == "cancelled":
            initial["status"] = "cancelled"
            initial["failure"] = evidence["failure"]
        if evidence["status"] != "memory_limit" or cell_budget == 1:
            break
        cell_budget = max(1, cell_budget // 4)
    if len(attempts) > 1:
        initial.setdefault(
            "recovery", {"status": attempts[-1]["status"], "mode": "source_only"}
        )
        initial["attempts"] = attempts
        initial["elapsedSeconds"] = round(time.monotonic() - started, 3)
    return result


def _run_isolated_once(
    data: bytes,
    kwargs: dict,
    policy: ExecutionPolicy,
    *,
    recovery_cells: int | None = None,
) -> dict:
    """Run one worker, enforcing elapsed time even inside a native call."""
    started = time.monotonic()
    status = "timed_out"
    checkpoint: dict = {}
    payload: dict = {}
    exit_code = None
    diagnostic = ""
    diagnostic_truncated = False
    with _worker_directory() as directory:
        root = Path(directory)
        (root / "input.xlsx").write_bytes(data)
        (root / "request.json").write_text(
            json.dumps(
                {
                    "kwargs": kwargs,
                    "memory_mb": policy.memory_mb,
                    "recovery_cell_budget": recovery_cells,
                },
                default=str,
            ),
            encoding="utf-8",
        )
        if policy.seconds > time.monotonic() - started:
            with (root / "stderr.txt").open("w", encoding="utf-8") as log:
                command = [
                    sys.executable,
                    "-X",
                    "faulthandler",
                    "-m",
                    "linexcel._worker",
                    str(root),
                ]
                if os.name != "nt":
                    # Apply the virtual-address limit before importing linexcel.
                    bootstrap = (
                        "import resource,runpy,sys; "
                        "resource.setrlimit(resource.RLIMIT_AS,"
                        f"({policy.memory_mb * 1048576},)*2); "
                        "sys.argv=['linexcel._worker',sys.argv[1]]; "
                        "runpy.run_module('linexcel._worker',run_name='__main__')"
                    )
                    command = [
                        sys.executable,
                        "-X",
                        "faulthandler",
                        "-c",
                        bootstrap,
                        str(root),
                    ]
                environment = os.environ.copy()
                environment.setdefault("RUST_BACKTRACE", "1")
                process = subprocess.Popen(
                    command,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=log,
                    env=environment,
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
                        size = (root / "stderr.txt").stat().st_size
                        diagnostic_truncated = size > DIAGNOSTIC_BYTES
                        if diagnostic_truncated:
                            # Keep the panic/allocation headline as well as the
                            # end of the trace; long Rust backtraces hide the
                            # actual error if only their tail survives.
                            head = errors.read(DIAGNOSTIC_BYTES // 2)
                            errors.seek(size - DIAGNOSTIC_BYTES // 2)
                            raw = (
                                head
                                + b"\n[... diagnostic truncated ...]\n"
                                + errors.read()
                            )
                        else:
                            raw = errors.read()
                        diagnostic = raw.decode("utf-8", errors="replace")
        failure = _failure_details(status, exit_code, diagnostic)
        if status == "crashed" and failure["kind"] == "memory_limit":
            status = "memory_limit"
        if status == "error" and recovery_cells is None:
            raise ValueError(payload.get("error", "Analysis worker failed"))
        if status == "error":
            failure["summary"] = str(payload.get("error", "Source recovery failed"))[
                :1024
            ]
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
                        "warnings": list(checkpoint.get("warnings", [])),
                    },
                    "nodes": [],
                    "edges": [],
                    "sheets": checkpoint.get("sheets", []),
                },
            )
            context = ""
            if checkpoint.get("operation"):
                context += f"Operation: {checkpoint['operation']}. "
            if exit_code is not None:
                context += f"Worker exit: {failure.get('exitCodeHex', exit_code)}. "
            if failure["summary"]:
                context += failure["summary"] + " "
            graph["meta"].setdefault("warnings", []).append(
                f"Analysis {status} during {checkpoint.get('phase', 'startup')}. "
                + context
                + "The requested analysis is incomplete; "
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
            "diagnosticTruncated": diagnostic_truncated,
            "operation": checkpoint.get("operation"),
            "failure": failure if status != "completed" else None,
            "versions": _runtime_versions(),
        }
        add_coverage(graph)
        return {"graph": graph, "engine": None, "analysisId": uuid.uuid4().hex[:16]}
