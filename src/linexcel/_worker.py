"""Private subprocess entry point, safe from notebook spawn recursion."""

from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path
from typing import Any


def atomic_json(path: Path, value: dict) -> None:
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    temporary.replace(path)


def main() -> None:
    root = Path(sys.argv[1])
    request = json.loads((root / "request.json").read_text(encoding="utf-8"))
    # Parent attaches the Windows Job before opening this gate.
    deadline = time.monotonic() + 10
    while not (root / "ready").exists():
        if time.monotonic() >= deadline:
            return
        time.sleep(0.01)
    if os.name != "nt":
        import resource

        limit = request["memory_mb"] * 1024 * 1024
        resource.setrlimit(resource.RLIMIT_AS, (limit, limit))
    from linexcel import progress
    from linexcel.analyzer import analyze_workbook
    from linexcel.execution import ExecutionPolicy

    state: dict[str, Any] = {
        "phase": "analysis",
        "completedPhases": [],
        "phaseMetrics": {},
    }
    phase_starts = {}
    checkpoint_index = 0

    def checkpoint(
        phase: str, completed: bool = False, evidence: dict | None = None
    ) -> None:
        nonlocal checkpoint_index
        if phase != state["phase"]:
            state.pop("operation", None)
        state["phase"] = phase
        if completed:
            if phase not in state["completedPhases"]:
                state["completedPhases"].append(phase)
            if phase in phase_starts:
                state["phaseMetrics"][phase] = {
                    "elapsedSeconds": round(
                        time.monotonic() - phase_starts.pop(phase), 4
                    )
                }
        else:
            phase_starts.setdefault(phase, time.monotonic())
        if evidence:
            state.update(evidence)
        # Immutable files avoid Windows replace/read sharing races.
        atomic_json(root / f"checkpoint-{checkpoint_index:06d}.json", state)
        checkpoint_index += 1

    progress._observer = checkpoint
    checkpoint("structure")
    try:
        result = analyze_workbook(
            (root / "input.xlsx").read_bytes(),
            execution=ExecutionPolicy(isolated=False),
            **request["kwargs"],
        )
        checkpoint("serialization")
        output = {"status": "completed", "graph": result["graph"]}
    except MemoryError:
        output = {"status": "memory_limit"}
    except Exception as exc:
        output = {"status": "error", "error": str(exc)}
    atomic_json(root / "result.json", output)


if __name__ == "__main__":
    main()
