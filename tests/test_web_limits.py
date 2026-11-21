"""Web limits: machine-derived defaults, per-operation budgets, named failures."""

from __future__ import annotations

import json
import os
import re
import threading

import pytest
from test_execution import workload
from test_web import request, wait, workbook

from linexcel.web import create_app

GiB = 1024**3
MiB = 1024**2


@pytest.mark.parametrize(
    ("physical", "expected"),
    [
        (None, 2048),
        (4096, 2048),
        (32768, 16384),
        # A "32 GB" machine reports a little less: half of it rounds down to 15 GiB.
        (32510, 15360),
        (12000, 5120),
    ],
)
def test_launcher_memory_is_half_the_machine_in_whole_gib_never_below_2_gib(
    physical, expected
):
    from linexcel.web_limits import launcher_memory_mb

    assert launcher_memory_mb(physical) == expected


@pytest.mark.parametrize(
    ("used", "free", "total", "expected_mb"),
    [
        # Small disk: keep a tenth of it free.
        (0, 10 * GiB, 20 * GiB, 8 * 1024),
        # Nearly full disk: the old 2 GiB of room, as long as the disk has it,
        # with the data directory's 512 MiB rounded away.
        (512 * MiB, 5 * GiB, 100 * GiB, 2048),
        (512 * MiB, 1 * GiB, 100 * GiB, 1024),
        # Large disk: the reserve stops at 10 GiB.
        (1 * GiB, 500 * GiB, 1000 * GiB, 1024 + 490 * 1024),
        # 300 MiB held plus 128 GiB 700 MiB of room: 128 GiB.
        (300 * MiB, 138 * GiB + 700 * MiB, 1000 * GiB, 128 * 1024),
        # Under 1 GiB of room, every MiB counts.
        (0, 600 * MiB, 100 * GiB, 600),
    ],
)
def test_launcher_storage_is_free_space_minus_a_reserve_in_whole_gib(
    used, free, total, expected_mb
):
    from linexcel.web_limits import launcher_storage_mb

    assert launcher_storage_mb(used, free, total) == expected_mb


@pytest.mark.parametrize(
    ("total", "workers", "expected"),
    [
        (15360, 2, 7680),
        # 9300 / 2 = 4650: down to the 512 MiB step.
        (9300, 2, 4608),
        (2048, 4, 1024),
        # Never more than the server has.
        (800, 1, 800),
    ],
)
def test_operation_default_memory_is_an_equal_share_in_512_mib_steps(
    total, workers, expected
):
    from linexcel.web_limits import TaskDefaults

    assert TaskDefaults.for_server(total, workers).memory_mb == expected


def _config(**limits):
    from pathlib import Path

    from linexcel.web import WebConfig

    return WebConfig(data_dir=Path("unused"), **limits)


KEPT = "Le classeur et les autres résultats restent accessibles."


@pytest.mark.parametrize(
    ("operation", "seconds", "expected"),
    [
        (
            "import",
            90,
            "Durée maximale de 90 s atteinte. Augmentez « Opérations longues » dans "
            "Réglages › Limites par opération (au plus 60 min), puis relancez. " + KEPT,
        ),
        (
            "capture",
            3600,
            "Durée maximale de 3600 s atteinte, la plus longue permise. " + KEPT,
        ),
        (
            "evaluate",
            45,
            "Durée maximale de 45 s atteinte. Augmentez « Recalcul ciblé » dans "
            "Réglages › Limites par opération (au plus 3600 s), puis relancez. " + KEPT,
        ),
        (
            "document",
            600,
            "Durée maximale de 600 s atteinte. Augmentez « Durée maximale » dans "
            "Réglages › Documentation IA (au plus 3600 s), puis relancez. " + KEPT,
        ),
    ],
)
def test_time_limit_names_the_value_and_the_field_of_the_settings(
    operation, seconds, expected
):
    from linexcel.web_limits import limit_message

    message = limit_message("timeout", _config(), operation=operation, seconds=seconds)
    assert message == expected


@pytest.mark.parametrize(
    ("memory_mb", "ceiling", "expected"),
    [
        (
            4608,
            15360,
            "Budget mémoire de 4,5 Go (4608 Mio) atteint. Augmentez « Mémoire par "
            "opération » dans Réglages › Limites par opération, jusqu’au plafond du "
            "serveur de 15 Go, ou relancez le serveur avec une valeur --memory-mb "
            "plus grande que 15360. " + KEPT,
        ),
        (
            2048,
            2048,
            "Budget mémoire de 2 Go (2048 Mio) atteint, égal au plafond du serveur. "
            "Relancez le serveur avec une valeur --memory-mb plus grande que 2048. "
            + KEPT,
        ),
    ],
)
def test_memory_limit_names_the_budget_the_ceiling_and_the_remedy(
    memory_mb, ceiling, expected
):
    from linexcel.web_limits import limit_message

    message = limit_message(
        "memory_limit",
        _config(total_memory_mb=ceiling),
        operation="import",
        memory_mb=memory_mb,
    )
    assert message == expected


def _storage_message(consequence):
    return (
        "Quota de stockage du serveur atteint (3 Mio, classeurs, index et vues "
        f"compris) : {consequence} Supprimez des classeurs ou relancez le serveur "
        "avec une valeur --max-storage-mb plus grande."
    )


def test_this_machine_reports_its_physical_memory():
    from linexcel.web_limits import physical_memory_mb

    memory = physical_memory_mb()
    if os.name == "nt" or os.path.exists("/proc/meminfo"):
        assert isinstance(memory, int) and memory >= 512


@pytest.fixture
def large_app(tmp_path):
    app = create_app(
        tmp_path / "workspace",
        total_memory_mb=8192,
        max_workers=2,
        max_storage_mb=4096,
        resolve_user=lambda env: env.get("user"),
    )
    yield app
    app.close()


def test_config_reports_quota_usage_and_operation_defaults(large_app):
    code, body, _ = request(large_app, "GET", "/api/config")
    assert code == 200
    limits = json.loads(body)["limits"]
    assert limits["memoryMb"] == 8192
    assert limits["storageMb"] == 4096
    assert isinstance(limits["storageUsedMb"], int) and limits["storageUsedMb"] >= 0
    assert limits["maxSeconds"] == 3600
    assert limits["defaults"] == {
        "longSeconds": 3600,
        "seconds": 120,
        # Each of the two workers gets half of the 8 GiB server budget.
        "memoryMb": 4096,
    }


def test_tasks_without_a_budget_get_the_defaults_of_their_operation(
    large_app, monkeypatch
):
    code, body, _ = request(
        large_app, "POST", "/api/projects", {"workbook": workbook()}
    )
    assert code == 200
    imported = json.loads(body)
    assert imported["task"]["budget"]["seconds"] == 3600
    assert imported["task"]["budget"]["memoryMb"] == 4096
    project = imported["project"]["id"]
    assert wait(large_app.store, imported["task"]["id"])["status"] == "succeeded"
    monkeypatch.setattr(
        large_app.store,
        "_worker",
        lambda task, root, event: {"error": {"kind": "stub", "message": "stub"}},
    )
    budgets = {}
    for operation, node in [("evaluate", "Sheet!B1"), ("capture", None)]:
        code, body, _ = request(
            large_app,
            "POST",
            f"/api/projects/{project}/tasks",
            {"operation": operation, "nodeId": node},
        )
        assert code == 200, body
        task_id = json.loads(body)["task"]["id"]
        _, body, _ = request(large_app, "GET", f"/api/tasks/{task_id}")
        budgets[operation] = json.loads(body)["task"]["budget"]
    assert budgets["evaluate"]["seconds"] == 120
    assert budgets["capture"]["seconds"] == 3600
    assert {budget["memoryMb"] for budget in budgets.values()} == {4096}


def _task_root(app, name="task"):
    root = app.store.root / "scratch" / name
    root.mkdir(parents=True)
    (root / "ready").touch()  # The workload child starts at once.
    return root


def _import(seconds=30, memory_mb=512):
    return {
        "operation": "import",
        "nodeId": None,
        "budget": {"memoryMb": memory_mb, "seconds": seconds},
    }


def test_time_limit_names_the_duration_and_the_setting(tmp_path, monkeypatch):
    app = create_app(tmp_path / "workspace")
    try:
        children = workload(monkeypatch, "time.sleep(30)")
        error = app.store._worker(
            _import(seconds=1), _task_root(app), threading.Event()
        )["error"]
        assert error["kind"] == "timeout"
        assert re.search(r"\b1 s\b", error["message"])
        assert "« Opérations longues »" in error["message"]
        assert all(child.poll() is not None for child in children)
    finally:
        app.close()


@pytest.mark.parametrize("child_exits", [False, True])
def test_storage_quota_names_the_quota_and_the_flag(tmp_path, monkeypatch, child_exits):
    app = create_app(tmp_path / "workspace", max_storage_mb=3)
    try:
        finish = (
            "(root/'result.json').write_text(json.dumps({'result':{}}))"
            if child_exits
            else "time.sleep(30)"
        )
        workload(
            monkeypatch,
            f"(root/'filler.bin').write_bytes(bytes(4*1048576))\n{finish}",
        )
        error = app.store._worker(_import(), _task_root(app), threading.Event())[
            "error"
        ]
        assert error["kind"] == "storage_limit"
        assert error["message"] == _storage_message(
            "l’opération a été arrêtée et son résultat n’a pas été conservé."
        )
    finally:
        app.close()


def test_storage_refusal_names_the_quota_and_the_flag(tmp_path):
    app = create_app(
        tmp_path / "workspace",
        max_storage_mb=3,
        resolve_user=lambda env: env.get("user"),
    )
    try:
        (app.store.root / "filler.bin").write_bytes(bytes(4 * MiB))
        code, body, _ = request(app, "POST", "/api/projects", {"workbook": workbook()})
        assert code == 413
        assert json.loads(body)["error"]["message"] == _storage_message(
            "l’opération n’a pas pu démarrer."
        )
    finally:
        app.close()


def test_storage_refusal_of_an_upload_chunk_says_the_transfer_stopped(tmp_path):
    from test_web_upload import begin, put

    app = create_app(
        tmp_path / "workspace",
        max_storage_mb=3,
        resolve_user=lambda env: env.get("user"),
    )
    try:
        upload, raw = begin(app)
        (app.store.root / "filler.bin").write_bytes(bytes(4 * MiB))
        code, body, _ = put(app, upload, raw)
        assert code == 413
        assert json.loads(body)["error"]["message"] == _storage_message(
            "le transfert a été interrompu."
        )
    finally:
        app.close()


def test_child_memory_error_names_the_budget_ceiling_and_flag(tmp_path, monkeypatch):
    # Linux reports RLIMIT_AS exhaustion as a Python MemoryError with no text.
    app = create_app(tmp_path / "workspace", total_memory_mb=2048)
    try:
        workload(
            monkeypatch,
            "(root/'result.json').write_text(json.dumps("
            "{'error':{'kind':'MemoryError','message':''}}))",
        )
        error = app.store._worker(
            _import(memory_mb=512), _task_root(app), threading.Event()
        )["error"]
        assert error["kind"] == "memory_limit"
        assert "0,5 Go (512 Mio)" in error["message"]
        assert "plafond du serveur de 2 Go" in error["message"]
        assert "« Mémoire par opération »" in error["message"]
        assert "--memory-mb plus grande que 2048" in error["message"]
    finally:
        app.close()


def test_supervision_does_not_scan_storage_on_every_pass(tmp_path, monkeypatch):
    app = create_app(tmp_path / "workspace")
    try:
        scan = app.store._storage_bytes
        scans = []

        def counted():
            scans.append(1)
            return scan()

        monkeypatch.setattr(app.store, "_storage_bytes", counted)
        workload(
            monkeypatch,
            "time.sleep(1.5)\n"
            "(root/'result.json').write_text(json.dumps({'result':{'done':True}}))",
        )
        result = app.store._worker(_import(), _task_root(app), threading.Event())
        assert result == {"result": {"done": True}}
        # One scan when supervision starts, one after a second, one exact check
        # before publishing; a scan per 50 ms pass would be about thirty.
        assert len(scans) <= 4, len(scans)
    finally:
        app.close()


GRANDCHILD = (
    "import subprocess\nr=subprocess.run([sys.executable,'-c',{code!r}])\n{finish}"
)


@pytest.mark.skipif(os.name != "nt", reason="Windows Job Object memory cap")
@pytest.mark.parametrize("reported", [False, True])
@pytest.mark.parametrize(
    ("memory_mb", "grandchild", "expected"),
    [
        # A descendant (like LibreOffice) dies of the job cap with exit code 1.
        (128, "bytearray(512*1048576)", "memory_limit"),
        # A descendant fails for another reason under a roomy cap.
        (1024, "import sys; sys.exit(3)", None),
    ],
)
def test_windows_job_memory_cap_wins_over_the_exit_code(
    tmp_path, monkeypatch, reported, memory_mb, grandchild, expected
):
    app = create_app(tmp_path / "workspace", total_memory_mb=2048)
    try:
        finish = (
            "(root/'result.json').write_text(json.dumps({'error':{"
            "'kind':'stack_overflow','message':'LibreOffice failed'}}))"
            if reported
            else "sys.exit(r.returncode)"
        )
        workload(monkeypatch, GRANDCHILD.format(code=grandchild, finish=finish))
        error = app.store._worker(
            _import(memory_mb=memory_mb), _task_root(app), threading.Event()
        )["error"]
        if expected == "memory_limit":
            assert error["kind"] == "memory_limit"
            assert re.search(rf"\b{memory_mb} Mio\b", error["message"])
            assert "--memory-mb" in error["message"]
        else:
            assert error["kind"] == ("stack_overflow" if reported else "unknown_exit")
        assert error["memoryLimitKind"] == "job_commit"
        assert error["memoryPeakMiB"] > 0
    finally:
        app.close()


@pytest.mark.parametrize("path", ["chunked upload", "single request"])
def test_an_import_over_the_size_limit_names_the_limit_and_the_flag(tmp_path, path):
    import base64

    app = create_app(
        tmp_path / "workspace",
        max_upload_mb=1,
        resolve_user=lambda env: env.get("user"),
    )
    try:
        if path == "chunked upload":
            body = {"workbook": {"name": "big.xlsx", "size": 2 * MiB}}
            code, raw, _ = request(app, "POST", "/api/uploads", body)
        else:
            data = base64.b64decode(workbook()["data"]) + bytes(2 * MiB)
            body = {
                "workbook": {
                    "name": "big.xlsx",
                    "data": base64.b64encode(data).decode(),
                }
            }
            code, raw, _ = request(app, "POST", "/api/projects", body)
        assert code == 413
        message = json.loads(raw)["error"]["message"]
        assert "1 Mio" in message
        assert "--max-upload-mb" in message
    finally:
        app.close()
