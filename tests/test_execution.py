"""Real subprocess tests: monkeypatch only worker workload, not supervision."""

import os
import subprocess
import sys
import time

import pytest

from linexcel import ExecutionPolicy, analyze
from linexcel.execution import add_coverage, run_isolated


def workload(monkeypatch, code):
    real_popen = subprocess.Popen
    children = []

    def launch(args, **kwargs):
        # Use the real interpreter, process gate and parent Job/session.
        root = args[-1]
        script = (
            "import sys,time,json,os; from pathlib import Path; "
            "root=Path(sys.argv[1]);\n"
            "while not (root/'ready').exists(): time.sleep(.01)\n" + code
        )
        child = real_popen([sys.executable, "-c", script, root], **kwargs)
        children.append(child)
        return child

    monkeypatch.setattr(subprocess, "Popen", launch)
    return children


def test_default_isolated_and_explicit_live_engine(lineage_excel):
    isolated = analyze(lineage_excel)
    assert isolated.engine is None
    assert isolated.graph["meta"]["execution"]["status"] == "completed"
    live = analyze(lineage_excel, execution=ExecutionPolicy(isolated=False))
    assert live.engine is not None
    assert live.nodes == isolated.nodes
    assert live.edges == isolated.edges


@pytest.mark.parametrize("seconds", [-1, float("inf"), float("nan"), True])
def test_invalid_budget_before_any_work(seconds):
    with pytest.raises((ValueError, TypeError)):
        ExecutionPolicy(seconds=seconds)


def test_zero_budget_never_starts_worker(monkeypatch):
    def fail(*args, **kwargs):
        pytest.fail("zero budget must not start a native worker")

    monkeypatch.setattr(subprocess, "Popen", fail)
    result = run_isolated(b"", {}, ExecutionPolicy(seconds=0))
    assert result["graph"]["meta"]["execution"]["status"] == "timed_out"
    assert result["graph"]["nodes"] == []


def test_timeout_interrupts_uncooperative_process(monkeypatch):
    children = workload(monkeypatch, "time.sleep(30)")
    start = time.monotonic()
    result = run_isolated(b"", {}, ExecutionPolicy(seconds=0.5))
    assert time.monotonic() - start < 5
    assert result["graph"]["meta"]["execution"]["status"] == "timed_out"
    assert all(child.poll() is not None for child in children)


def test_native_exit_does_not_kill_caller_or_publish_values(monkeypatch):
    children = workload(monkeypatch, "os._exit(17)")
    result = run_isolated(b"", {}, ExecutionPolicy(seconds=5))
    assert result["graph"]["meta"]["execution"]["status"] == "crashed"
    assert result["graph"]["nodes"] == []
    assert children[0].returncode == 17


def test_cancellation_cleans_worker(monkeypatch):
    children = workload(monkeypatch, "time.sleep(30)")
    real_sleep = time.sleep
    interruptions = 0

    def interrupt(seconds):
        nonlocal interruptions
        if interruptions == 0:
            interruptions += 1
            raise KeyboardInterrupt
        # POSIX subprocess.wait also sleeps while cleanup reaps the worker.
        # Inject one cancellation, not another interruption of that cleanup.
        real_sleep(seconds)

    monkeypatch.setattr(time, "sleep", interrupt)
    result = run_isolated(b"", {}, ExecutionPolicy(seconds=5))
    assert result["graph"]["meta"]["execution"]["status"] == "cancelled"
    assert interruptions == 1
    assert len(children) == 1
    assert all(child.poll() is not None for child in children)


@pytest.mark.skipif(os.name != "nt", reason="Windows Job aggregate memory contract")
def test_windows_memory_cap_is_effective(monkeypatch):
    workload(
        monkeypatch,
        """
try:
    values=[]
    for _ in range(256): values.append(bytearray(1024*1024))
    status='completed'
except MemoryError:
    status='memory_limit'
(root/'result.json').write_text(json.dumps({'status':status}))
""",
    )
    result = run_isolated(b"", {}, ExecutionPolicy(seconds=5, memory_mb=64))
    assert result["graph"]["meta"]["execution"]["status"] == "memory_limit"


def test_no_descendant_survives_success(monkeypatch, tmp_path):
    marker = tmp_path / "survived.txt"
    child_script = (
        "import time; from pathlib import Path; time.sleep(1); "
        f"Path({str(marker)!r}).touch()"
    )
    workload(
        monkeypatch,
        f"""
import subprocess
subprocess.Popen([sys.executable,'-c',{child_script!r}])
(root/'result.json').write_text(json.dumps({{'status':'completed','graph':{{'meta':{{}},'nodes':[],'edges':[]}}}}))
""",
    )
    run_isolated(b"", {}, ExecutionPolicy(seconds=5))
    time.sleep(1.2)
    assert not marker.exists()


def test_coverage_counts_nodes_without_extrapolating_groups():
    graph = {
        "meta": {},
        "nodes": [
            {
                "id": "group",
                "kind": "group",
                "count": 500,
                "value": 3,
                "valueSource": "engine",
                "cachedAgreement": "differ",
            },
            {"id": "file", "value": 4, "valueSource": "file"},
            {"id": "missing", "kind": "cell", "value": None},
        ],
    }
    add_coverage(graph)
    coverage = graph["meta"]["coverage"]
    assert coverage["totalNodes"] == 3
    assert coverage["categories"]["engine"]["count"] == 1
    assert coverage["categories"]["divergent"]["nodeIds"] == ["group"]


def test_range_samples_and_constants_are_not_missing_or_recalculated():
    graph = {
        "meta": {},
        "nodes": [
            {"id": "range", "kind": "input", "values": [{"value": 3}]},
            {"id": "constant", "kind": "input", "value": 3, "valueSource": "engine"},
            {"id": "vba", "kind": "vba"},
            {
                "id": "group",
                "kind": "group",
                "value": 3,
                "valueSource": "engine",
                "cachedAgreement": "same",
                "groupCachedAgreement": "differ",
            },
        ],
    }
    add_coverage(graph)
    categories = graph["meta"]["coverage"]["categories"]
    assert categories["other"]["count"] == 3
    assert categories["unavailable"]["count"] == 0
    assert categories["engine"]["count"] == 1
    assert categories["divergent"]["nodeIds"] == ["group"]


def test_timeout_overview_preserves_unknown_counts(lineage_excel):
    from linexcel.aidoc import build_workbook_dossier

    result = analyze(lineage_excel, execution=ExecutionPolicy(seconds=0))
    dossier = build_workbook_dossier(
        result.graph, context={"sheets": [{"name": "Sales"}]}
    )
    assert dossier["analysis"]["formula_cells"] is None
    assert dossier["sheets"][0]["formula_cells"] is None
    assert dossier["execution"]["status"] == "timed_out"


def test_cancelled_cli_does_not_start_optional_work(monkeypatch, tmp_path):
    import linexcel
    from linexcel.cli import main

    graph = {"meta": {"execution": {"status": "cancelled"}}, "nodes": [], "edges": []}
    monkeypatch.setattr(
        linexcel, "analyze", lambda *a, **k: linexcel.LineageResult(graph, None)
    )

    def fail(*a, **k):
        pytest.fail("cancellation must not start another stage")

    monkeypatch.setattr(linexcel.LineageResult, "document", fail)
    monkeypatch.setattr(linexcel.LineageResult, "save_screenshots", fail)
    assert (
        main(
            [
                "analyze",
                str(tmp_path / "book.xlsx"),
                "--ai-docs",
                "--screenshots",
                str(tmp_path / "shots"),
            ]
        )
        == 130
    )
