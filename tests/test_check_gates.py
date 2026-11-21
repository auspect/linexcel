"""CI and the commit hook run the same gates, through scripts/check.py only."""

import importlib.util
import re
import tomllib
from pathlib import Path

import pytest
import yaml

ROOT = Path(__file__).parents[1]
WORKFLOWS = sorted((ROOT / ".github" / "workflows").glob("*.yml"))
SCRIPT = "scripts/check.py"


@pytest.fixture(scope="module")
def gates():
    spec = importlib.util.spec_from_file_location("check_gates", ROOT / SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.GATES


def workflow(path: Path) -> dict:
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    # YAML 1.1 reads the bare key `on` as the boolean true.
    data["on"] = data.pop(True, data.get("on"))
    return data


def workflow_commands(path: Path) -> list[str]:
    return [
        step["run"]
        for job in workflow(path)["jobs"].values()
        for step in job.get("steps", [])
        if "run" in step
    ]


def prek_hooks() -> list[dict]:
    config = tomllib.loads((ROOT / "prek.toml").read_text(encoding="utf-8"))
    return [hook for repo in config["repos"] for hook in repo["hooks"]]


def commands() -> list[tuple[str, str]]:
    found = [(path.name, run) for path in WORKFLOWS for run in workflow_commands(path)]
    return found + [("prek.toml", hook["entry"]) for hook in prek_hooks()]


def lines(command: str) -> list[tuple[list[str], list[str]]]:
    """Per line, the words that run tools and the gate names given to the script."""
    split = []
    for line in command.splitlines():
        words = re.findall(r"[^\s'\"]+", line)
        at = words.index(SCRIPT) if SCRIPT in words else len(words)
        split.append((words[:at], words[at + 1 :]))
    return split


def gates_called(command: str) -> list[str]:
    return [name for _, names in lines(command) for name in names]


def signature(argv: tuple[str, ...]) -> list[str]:
    program, first, *_ = argv
    return [first] if program == "python" else [program, first]


def runs(command: str, sequence: list[str]) -> bool:
    return any(
        words[i : i + len(sequence)] == sequence
        for words, _ in lines(command)
        for i in range(len(words) - len(sequence) + 1)
    )


def test_ci_and_hook_run_gate_tools_only_through_the_script(gates):
    direct = [
        f"{source}: `{command.strip()}` runs `{' '.join(signature(gate.argv))}`"
        f" directly; call `uv run python {SCRIPT} {name}` instead"
        for source, command in commands()
        for name, gate in gates.items()
        if runs(command, signature(gate.argv))
    ]
    assert not direct, (
        f"Gate commands are defined once, in {SCRIPT}, so CI and the hook cannot"
        " drift apart:\n" + "\n".join(direct)
    )


def test_every_gate_named_in_ci_and_hook_exists(gates):
    unknown = [
        f"{source}: {name}"
        for source, command in commands()
        for name in gates_called(command)
        if name not in gates
    ]
    assert not unknown, f"Not a gate of {SCRIPT}: {unknown}"


def test_gate_hooks_run_on_every_commit():
    """Only a gate knows which files it reads; a hook pattern copy goes stale."""
    filtered = [
        hook["id"]
        for hook in prek_hooks()
        if gates_called(hook["entry"])
        and (
            not hook.get("always_run")
            or hook.get("pass_filenames", True)
            or hook.keys() & {"files", "exclude", "types", "types_or", "exclude_types"}
        )
    ]
    assert not filtered, (
        f"prek.toml hooks {filtered} filter the files that fire them. A gate of"
        f" {SCRIPT} decides its own inputs: set always_run = true and"
        " pass_filenames = false, with no files or types filter"
    )


def test_ci_runs_every_gate(gates):
    in_ci = {
        name
        for command in workflow_commands(ROOT / ".github" / "workflows" / "ci.yml")
        for name in gates_called(command)
    }
    assert in_ci == set(gates), (
        f"ci.yml must run every gate of {SCRIPT}; missing {set(gates) - in_ci}"
    )


@pytest.mark.parametrize("path", WORKFLOWS, ids=lambda path: path.name)
def test_pull_requests_are_checked_whatever_their_base(path):
    trigger = workflow(path)["on"]
    pull_request = trigger.get("pull_request") if isinstance(trigger, dict) else None
    assert not (pull_request or {}).keys() & {"branches", "branches-ignore"}, (
        f"{path.name}: a branch filter skips stacked pull requests; filter by"
        " `paths` or guard a deployment job with `if:` instead"
    )
