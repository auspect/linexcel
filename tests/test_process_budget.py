"""Portable startup and real supervision of the macOS memory backend."""

import ctypes
import errno
import os
import sys
import threading
from types import SimpleNamespace

import pytest
from test_execution import workload

from linexcel import ExecutionPolicy
from linexcel import _process_budget as budget
from linexcel.execution import run_isolated
from linexcel.web import create_app


def test_startup_uses_the_platform_memory_boundary(tmp_path):
    command = budget.worker_command("linexcel._worker", tmp_path, 256)
    assert command[-1] == str(tmp_path)
    if os.name == "nt" or sys.platform == "darwin":
        assert command[3:5] == ["-m", "linexcel._worker"]
    else:
        assert command[3] == "-c"
        assert "resource.setrlimit(resource.RLIMIT_AS,(268435456,)*2)" in command[4]
        assert command[4].index("setrlimit") < command[4].index("runpy.run_module")


@pytest.mark.skipif(os.name == "nt", reason="Darwin uses POSIX process groups")
def test_darwin_does_not_install_an_address_space_cap(monkeypatch, tmp_path):
    monkeypatch.setattr(sys, "platform", "darwin")
    for module in ("linexcel._worker", "linexcel.web_worker"):
        command = budget.worker_command(module, tmp_path, 128)
        assert command[3:5] == ["-m", module]
        assert all("setrlimit" not in argument for argument in command)


@pytest.fixture
def native_api(monkeypatch):
    state = {"pids": [11, 12, 13], "rss": {11: 2**20, 12: 2 * 2**20}}

    def list_pids(group, buffer, size):
        assert group == 11
        if state.get("list_error"):
            ctypes.set_errno(state["list_error"])
            return 0
        if buffer is not None:
            assert size >= len(state["pids"]) * ctypes.sizeof(ctypes.c_int)
            for i, pid in enumerate(state["pids"]):
                buffer[i] = pid
        return len(state["pids"])  # libproc returns PID counts, not bytes.

    def pid_rusage(pid, flavor, pointer):
        assert flavor == 0
        if pid not in state["rss"]:
            ctypes.set_errno(state.get("read_error", errno.ESRCH))
            return -1
        usage = ctypes.cast(pointer, ctypes.POINTER(budget._RUsageInfoV0)).contents
        usage.resident_size = state["rss"][pid]
        return 0

    monkeypatch.setattr(
        ctypes,
        "CDLL",
        lambda *a, **k: SimpleNamespace(
            proc_listpgrppids=list_pids, proc_pid_rusage=pid_rusage
        ),
    )
    return state


def test_native_group_reader_counts_descendants_and_tolerates_exits(native_api):
    assert budget._DarwinGroupRSS()(11) == 3 * 2**20
    native_api["pids"] = []
    assert budget._DarwinGroupRSS()(11) is None


@pytest.mark.parametrize("stage", ["list_error", "read_error"])
def test_native_monitor_errors_do_not_become_an_empty_sample(native_api, stage):
    native_api[stage] = errno.EPERM
    with pytest.raises(PermissionError, match="Cannot supervise macOS worker memory"):
        budget._DarwinGroupRSS()(11)


@pytest.mark.skipif(os.name == "nt", reason="Darwin uses POSIX process groups")
def test_sampled_limit_and_peak_are_reported_without_claiming_a_hard_cap(monkeypatch):
    monkeypatch.setattr(sys, "platform", "darwin")
    samples = iter([2**20, 2 * 2**20, None, 3 * 2**20])
    monkeypatch.setattr(budget, "_DarwinGroupRSS", lambda: lambda _: next(samples))
    monitor = budget.MemoryMonitor(2)
    assert [monitor.exceeded(11) for _ in range(4)] == [False, False, False, True]
    assert monitor.evidence() == {
        "memoryLimitKind": "process_group_sampled_rss",
        "memoryPeakMiB": 3.0,
    }


@pytest.mark.skipif(os.name == "nt", reason="Darwin uses POSIX process groups")
def test_analysis_stops_the_group_without_publishing_values(monkeypatch):
    monkeypatch.setattr(sys, "platform", "darwin")
    monkeypatch.setattr(budget, "_DarwinGroupRSS", lambda: lambda _: 65 * 2**20)
    children = workload(monkeypatch, "time.sleep(30)")
    result = run_isolated(b"", {}, ExecutionPolicy(seconds=5, memory_mb=64))
    evidence = result["graph"]["meta"]["execution"]
    assert evidence["status"] == "memory_limit"
    assert evidence["memoryPeakMiB"] == 65.0
    assert evidence["failure"]["kind"] == "memory_limit"
    assert "sampled limit" in evidence["failure"]["summary"]
    assert not result["graph"]["nodes"]
    assert all(child.poll() is not None for child in children)


@pytest.mark.skipif(os.name == "nt", reason="Darwin uses POSIX process groups")
def test_native_memory_failure_does_not_invent_a_sampled_limit_breach(monkeypatch):
    monkeypatch.setattr(sys, "platform", "darwin")
    monkeypatch.setattr(budget, "_DarwinGroupRSS", lambda: lambda _: 32 * 2**20)
    workload(
        monkeypatch,
        "(root/'result.json').write_text(json.dumps({'status':'memory_limit'}))",
    )
    evidence = run_isolated(
        b"", {}, ExecutionPolicy(seconds=5, memory_mb=64, memory_retries=0)
    )["graph"]["meta"]["execution"]
    assert evidence["status"] == "memory_limit"
    assert evidence["memoryPeakMiB"] == 32.0
    assert "sampled limit" not in evidence["failure"]["summary"]


@pytest.mark.skipif(os.name == "nt", reason="Darwin uses POSIX process groups")
def test_monitor_failure_stops_analysis_instead_of_disabling_the_limit(monkeypatch):
    def unreadable(_):
        raise PermissionError("Cannot supervise macOS worker memory")

    monkeypatch.setattr(sys, "platform", "darwin")
    monkeypatch.setattr(budget, "_DarwinGroupRSS", lambda: unreadable)
    children = workload(monkeypatch, "time.sleep(30)")
    with pytest.raises(PermissionError, match="Cannot supervise macOS worker memory"):
        run_isolated(b"", {}, ExecutionPolicy(seconds=5, memory_mb=64))
    assert all(child.poll() is not None for child in children)


@pytest.mark.skipif(os.name == "nt", reason="Darwin uses POSIX process groups")
def test_application_uses_the_same_sampled_memory_limit(monkeypatch, tmp_path):
    app = create_app(tmp_path / "workspace")
    try:
        monkeypatch.setattr(sys, "platform", "darwin")
        monkeypatch.setattr(budget, "_DarwinGroupRSS", lambda: lambda _: 129 * 2**20)
        children = workload(monkeypatch, "time.sleep(30)")
        root = tmp_path / "task"
        root.mkdir()
        result = app.store._worker(
            {
                "operation": "import",
                "nodeId": None,
                "budget": {"memoryMb": 128, "seconds": 5},
            },
            root,
            threading.Event(),
        )
        assert result["error"]["kind"] == "memory_limit"
        assert result["error"]["memoryLimitKind"] == "process_group_sampled_rss"
        assert result["error"]["memoryPeakMiB"] == 129.0
        assert all(child.poll() is not None for child in children)
    finally:
        app.close()


@pytest.mark.skipif(sys.platform != "darwin", reason="Real libproc enforcement")
@pytest.mark.parametrize("descendant", [False, True])
def test_real_macos_group_memory_is_bounded(monkeypatch, descendant):
    allocate = (
        "import time\nvalues=[]\n"
        "for _ in range(256):\n"
        " chunk=bytearray(1048576)\n"
        " chunk[::4096]=b'x'*256\n"
        " values.append(chunk)\n"
        " time.sleep(.01)\n"
        "time.sleep(30)\n"
    )
    code = (
        "import subprocess\n"
        f"subprocess.Popen([sys.executable,'-c',{allocate!r}])\n"
        "time.sleep(30)\n"
        if descendant
        else allocate
    )
    children = workload(monkeypatch, code)
    evidence = run_isolated(
        b"", {}, ExecutionPolicy(seconds=8, memory_mb=64, memory_retries=0)
    )["graph"]["meta"]["execution"]
    assert evidence["status"] == "memory_limit", evidence
    assert evidence["memoryPeakMiB"] > 64
    assert all(child.poll() is not None for child in children)
