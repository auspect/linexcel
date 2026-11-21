"""Every check that CI and the commit hook run, each defined once.

    uv run python scripts/check.py              # every gate
    uv run python scripts/check.py ty ruff      # the named gates

`.github/workflows/ci.yml` and `prek.toml` run these gates by name and never the
tools themselves (`tests/test_check_gates.py` fails otherwise), so a flag or a
rule renamed here changes the local run and CI together.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path
from typing import NamedTuple

ROOT = Path(__file__).resolve().parent.parent


class Gate(NamedTuple):
    argv: tuple[str, ...]
    fix: str
    passing: frozenset[int] = frozenset({0})


GATES: dict[str, Gate] = {
    "ruff": Gate(
        ("ruff", "check", "src/", "tests/"),
        "uv run ruff check --fix src/ tests/",
    ),
    "format": Gate(
        ("ruff", "format", "--check", "src/", "tests/"),
        "uv run ruff format src/ tests/",
    ),
    # The optional AI extras are not installed here, so their imports are unresolved.
    "ty": Gate(
        ("ty", "check", "--ignore", "import-not-found", "src/"),
        "fix the diagnostics above; rule names follow the ty in uv.lock",
    ),
    "viewer-captures": Gate(
        ("python", "scripts/capture_viewer.py", "--check"),
        "uv run python scripts/capture_viewer.py, then commit imgs/",
    ),
    "app-captures": Gate(
        ("python", "scripts/capture_app.py", "--check"),
        "recapture with scripts/capture_app.py (commands in its docstring),"
        " then commit imgs/",
    ),
    "frontend": Gate(
        ("npm", "run", "check", "--prefix", "frontend"),
        "npm ci --prefix frontend; then npm run build --prefix frontend"
        " and commit the bundle",
    ),
    # Exit 3: only the findings recorded in known_findings.json still fail.
    "lineage": Gate(
        (
            "python",
            ".agents/skills/check-group-precedents/scripts/check_group_precedents.py",
            "--quick",
        ),
        "read the counterexample above and .agents/skills/check-group-precedents",
        frozenset({0, 3}),
    ),
    "lineage-mutants": Gate(
        (
            "python",
            ".agents/skills/check-group-precedents/scripts/check_group_precedents.py",
            "--mutants",
        ),
        "a MISSED or STALE variant means the lineage check got weaker; see"
        " .agents/skills/check-group-precedents",
    ),
}


def run(name: str) -> bool:
    gate = GATES[name]
    program, *args = gate.argv
    executable = sys.executable if program == "python" else shutil.which(program)
    print(f"== {name}: {' '.join(gate.argv)}", flush=True)
    if executable is None:
        print(f"{program} is not on PATH: run `uv sync`, or install Node for npm")
        return False
    return subprocess.run([executable, *args], cwd=ROOT).returncode in gate.passing


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("gates", nargs="*", help=f"any of: {', '.join(GATES)}")
    names = parser.parse_args(argv).gates or list(GATES)
    unknown = [name for name in names if name not in GATES]
    if unknown:
        parser.error(f"unknown gate {', '.join(unknown)}; known: {', '.join(GATES)}")
    failed = [name for name in names if not run(name)]
    for name in failed:
        print(f"FAILED {name}. Fix: {GATES[name].fix}")
    if not failed:
        print(f"PASSED {', '.join(names)}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
