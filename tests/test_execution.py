"""Real subprocess tests: monkeypatch only worker workload, not supervision."""

import io
import json
import os
import subprocess
import sys
import time

import pytest

from linexcel import ExecutionPolicy, analyze
from linexcel.execution import _failure_details, add_coverage, run_isolated


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


def test_native_allocation_abort_is_reported_as_memory_failure(monkeypatch):
    workload(
        monkeypatch,
        """
(root/'checkpoint-000001.json').write_text(json.dumps({
    'phase':'engine evaluation', 'operation':'evaluating all formulas'}))
sys.stderr.write('memory allocation of 1048576 bytes failed\\n')
sys.stderr.flush()
os._exit(17)
""",
    )
    result = run_isolated(b"", {}, ExecutionPolicy(seconds=5))
    meta = result["graph"]["meta"]
    execution = meta["execution"]
    assert execution["status"] == "memory_limit"
    assert execution["failure"]["kind"] == "memory_limit"
    assert execution["operation"] == "evaluating all formulas"
    assert "Native memory allocation failed" in meta["warnings"][0]
    assert "evaluating all formulas" in meta["warnings"][0]
    assert execution["exitCode"] == 17
    assert execution["versions"]["formualizer"]
    assert not execution["diagnosticTruncated"]
    assert not result["graph"]["nodes"]


@pytest.mark.skipif(os.name != "nt", reason="Reproduced with Windows Job memory cap")
def test_real_engine_allocation_failure_keeps_native_evidence():
    from openpyxl import Workbook

    workbook = Workbook()
    workbook.active["A1"] = "=SUM(SEQUENCE(100000,100))"
    data = io.BytesIO()
    workbook.save(data)
    result = analyze(
        data.getvalue(), execution=ExecutionPolicy(seconds=30, memory_mb=128)
    )
    execution = result.graph["meta"]["execution"]
    assert execution["status"] == "memory_limit", execution
    assert execution["operation"] == "evaluating all formulas", execution
    assert "memory allocation of" in execution["diagnostic"], execution
    assert execution["failure"]["kind"] == "memory_limit"
    assert not result.nodes


@pytest.mark.parametrize(
    "stderr,kind",
    [
        ("thread '<unknown>' has overflowed its stack", "stack_overflow"),
        ("thread '<unnamed>' panicked at src/eval.rs:42:1:\nbad graph", "native_panic"),
        ("pyo3_runtime.PanicException: bad graph", "native_panic"),
        ("some formula contains memory allocation of 12 bytes failed", "unknown_exit"),
        ("", "unknown_exit"),
    ],
)
def test_crash_evidence_is_not_guessed(stderr, kind):
    assert _failure_details("crashed", 17, stderr)["kind"] == kind


@pytest.mark.skipif(os.name != "nt", reason="Windows exit codes")
@pytest.mark.parametrize(
    "code,kind",
    [
        (0xC00000FD, "stack_overflow"),
        (0xC0000005, "access_violation"),
        (0xC0000017, "memory_limit"),
        (0xC000012D, "memory_limit"),
        (0xC0000409, "native_abort"),
    ],
)
def test_windows_native_exit_codes(code, kind):
    for value in (code, code - 2**32):
        details = _failure_details("crashed", value, "")
        assert details["kind"] == kind
        assert details["exitCodeHex"] == f"0x{code:08X}"


@pytest.mark.skipif(os.name == "nt", reason="POSIX signals")
def test_sigkill_is_not_proof_of_oom():
    details = _failure_details("crashed", -9, "")
    assert details["signal"] == "SIGKILL"
    assert details["kind"] == "native_signal"


def test_diagnostic_tail_is_bounded_and_marked(monkeypatch):
    workload(
        monkeypatch, "sys.stderr.write('x'*20000); sys.stderr.flush(); os._exit(17)"
    )
    execution = run_isolated(b"", {}, ExecutionPolicy(seconds=5))["graph"]["meta"][
        "execution"
    ]
    assert execution["diagnosticTruncated"]
    assert execution["diagnostic"].count("x") == 16384
    assert "diagnostic truncated" in execution["diagnostic"]


@pytest.mark.parametrize(
    "helper,operation",
    [
        ("_find_too_deep", "checking formula complexity"),
        ("_open_workbook", "importing workbook into native engine"),
    ],
)
def test_real_worker_retains_operation_before_native_exit(
    monkeypatch, lineage_excel, helper, operation
):
    workload(
        monkeypatch,
        f"""
from linexcel import engine
from linexcel._worker import main
def crash(*args, **kwargs):
    sys.stderr.write('memory allocation of 1048576 bytes failed\\n' + 'x'*20000)
    sys.stderr.flush()
    os._exit(17)
engine.{helper} = crash
main()
""",
    )
    result = run_isolated(lineage_excel, {}, ExecutionPolicy(seconds=10))
    execution = result["graph"]["meta"]["execution"]
    assert execution["phase"] == "engine evaluation"
    assert execution["operation"] == operation
    assert execution["status"] == "memory_limit"
    assert execution["diagnosticTruncated"]
    assert "engine evaluation" not in execution["completedPhases"]
    assert not result["graph"]["nodes"]


def test_cli_diagnostics_excludes_graph_and_keeps_stdout_clean(
    monkeypatch, tmp_path, capsys
):
    import linexcel
    from linexcel.cli import main

    execution = {
        "status": "crashed",
        "exitCode": 17,
        "operation": "evaluating all formulas",
    }
    graph = {
        "meta": {"execution": execution, "warnings": []},
        "nodes": [{"id": "private", "kind": "cell", "secret": "private"}],
        "edges": [],
    }
    monkeypatch.setattr(
        linexcel, "analyze", lambda *a, **k: linexcel.LineageResult(graph, None)
    )
    path = tmp_path / "diagnostics.json"
    assert (
        main(["analyze", "private.xlsx", "--no-html", "--diagnostics", str(path)]) == 3
    )
    assert json.loads(path.read_text(encoding="utf-8")) == execution
    assert "private" not in path.read_text(encoding="utf-8")
    assert capsys.readouterr().out == ""


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
