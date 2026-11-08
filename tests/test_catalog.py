"""Deterministic library contracts: isolation, budgets and immutable generations."""

import hashlib
import json
import os
import sqlite3
import subprocess
import sys
import threading
from contextlib import closing

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table
from test_web_index import call

from linexcel.catalog import Catalog, CatalogLimits, IndexPool, Source, parse_handle
from linexcel.lazy_store import build_index
from linexcel.web import create_app


def source(tmp_path, number=1, *, rows=12, table=False, text="Chiffre d’affaires café"):
    directory = tmp_path / f"source{number}"
    directory.mkdir(exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.title = "Données"
    sheet.append(["Libellé", text])
    for row in range(2, rows + 1):
        sheet.append([f"Région {row}", row * number])
    sheet["C2"] = "=B2*2"
    if table:
        sheet.add_table(Table(displayName="Résultats", ref=f"A1:B{rows}"))
    book.save(directory / "workbook.xlsx")
    index = directory / "index.sqlite"
    build_index(directory / "workbook.xlsx", index, filename="same.xlsx")
    return Source(f"{number:032x}", "same.xlsx", index, 1)


def test_search_diacritics_numbers_source_provenance_and_handles(tmp_path):
    item = source(tmp_path)
    catalog = Catalog(tmp_path / "catalog")
    result = catalog.search([item], "CAFE", kinds=["cell"])
    assert result["coverage"]["complete"]
    assert result["trace"]["schemaVersion"] == "linexcel.search.v1"
    assert result["trace"]["terms"] == ["cafe"]
    assert result["trace"]["backend"] == "fts5"
    numeric = catalog.search([item], "4")
    assert [r["address"] for r in numeric["results"]] == ["B4"]
    hit = numeric["results"][0]
    assert hit["value"] == 4 and hit["valueProvenance"] == "saved_input"
    assert hit["labels"][0]["method"] == "left_text"
    assert hit["labels"][0]["cell"] == "A4"
    assert parse_handle(hit["handle"])[0] == item.project_id
    graph = catalog.lineage([item], hit["handle"])
    assert graph["revision"] == hit["revision"]
    assert not graph["coverage"]["complete"]
    assert "external_workbooks_not_traversed" in graph["coverage"]["boundaries"]
    with pytest.raises(KeyError):
        catalog.lineage([], hit["handle"])


def test_same_name_projects_are_distinct_and_scopes_are_checked(tmp_path):
    sources = [source(tmp_path, n) for n in (1, 2)]
    catalog = Catalog(tmp_path / "catalog")
    result = catalog.search(sources, "café", kinds=["cell"])
    assert {h["filename"] for h in result["results"]} == {"same.xlsx"}
    revisions = {h["revision"]["id"] for h in result["results"]}
    assert len(revisions) == 2
    scoped = catalog.search(sources, "cafe", revision_id=sorted(revisions)[0])
    assert len({r["revision"]["id"] for r in scoped["results"]}) == 1
    with pytest.raises(KeyError):
        catalog.search(sources, "", project="f" * 32)
    same_content = Source("f" * 32, sources[0].filename, sources[0].path)
    result = catalog.search([sources[0], same_content], "4")
    assert len(result["results"]) == 2
    identities = [hit["revision"] for hit in result["results"]]
    assert identities[0]["sha256"] == identities[1]["sha256"]
    assert identities[0]["id"] != identities[1]["id"]


def test_generation_atomicity_with_a_pinned_reader_and_concurrent_writer(tmp_path):
    first = source(tmp_path)
    second = source(tmp_path, 2)
    catalog = Catalog(tmp_path / "catalog")
    old = catalog.ensure([first])["generation"]
    errors = []
    with catalog.pin() as (db, current):
        count = db.execute("SELECT count(*) FROM entries").fetchone()[0]

        def build():
            try:
                catalog.ensure([first, second])
            except Exception as error:
                errors.append(error)

        worker = threading.Thread(target=build)
        worker.start()
        worker.join(10)
        assert not worker.is_alive() and not errors
        assert current["generation"] == old
        assert db.execute("SELECT count(*) FROM entries").fetchone()[0] == count
        assert (catalog.directory / f"{old}.sqlite").exists()
        assert catalog.status()["generation"] != old
    assert not (catalog.directory / f"{old}.sqlite").exists()
    assert not (catalog.directory / f"{old}.json").exists()
    with pytest.raises(ValueError, match="generation changed"):
        catalog.search([first, second], "", generation=old)


def test_content_hash_invalidates_even_when_mtime_is_restored_and_delete(tmp_path):
    item = source(tmp_path)
    catalog = Catalog(tmp_path / "catalog")
    first = catalog.search([item], "4")
    handle = first["results"][0]["handle"]
    stat = item.path.stat()
    # Replace in place with another genuine immutable index, then restore mtime.
    newer = source(tmp_path, 2)
    item.path.write_bytes(newer.path.read_bytes())
    os.utime(item.path, ns=(stat.st_atime_ns, stat.st_mtime_ns))
    second = catalog.search([item], "4")
    assert second["generation"] != first["generation"]
    assert second["results"][0]["revision"] != first["results"][0]["revision"]
    with pytest.raises(KeyError):
        catalog.lineage([item], handle)
    assert catalog.search([], "4")["results"] == []
    assert catalog.status()["revisions"] == []


def test_like_fallback_is_reported_normalized_and_bounded(tmp_path, monkeypatch):
    monkeypatch.setattr("linexcel.catalog.fts5_available", lambda: False)
    item = source(tmp_path)
    catalog = Catalog(tmp_path / "catalog")
    result = catalog.search([item], "café")
    assert result["results"] and result["trace"]["backend"] == "like"
    assert result["coverage"]["complete"]
    limited = Catalog(tmp_path / "limited", CatalogLimits(scan_entries=1))
    result = limited.search([item], "café")
    assert not result["coverage"]["complete"]
    assert "scan_entries" in result["trace"]["capsHit"]


@pytest.mark.parametrize(
    "limits,cap",
    [
        (CatalogLimits(entries_per_workbook=3), "entries_per_workbook"),
        (CatalogLimits(build_seconds=0), "build_seconds"),
        (CatalogLimits(hits=2), "hits"),
        (CatalogLimits(text_chars=2), "text_chars"),
        (CatalogLimits(query_ms=0), "query_ms"),
        (CatalogLimits(response_bytes=4096), "response_bytes"),
    ],
)
def test_caps_are_incomplete_with_counts(tmp_path, limits, cap):
    item = source(tmp_path, rows=30)
    catalog = Catalog(tmp_path / "catalog", limits)
    result = catalog.search([item], "", limit=min(50, limits.hits))
    assert cap in result["coverage"]["capsHit"]
    assert not result["coverage"]["complete"]
    assert result["coverage"]["nodesTotal"] == 61
    assert result["trace"]["counts"]["returned"] == len(result["results"])
    assert (
        len(json.dumps(result, ensure_ascii=False, separators=(",", ":")).encode())
        <= limits.response_bytes
    )


def test_disk_cap_publishes_searchable_prefix_and_counts(tmp_path):
    item = source(tmp_path, rows=120)
    catalog = Catalog(tmp_path / "catalog", CatalogLimits(catalog_bytes=65536))
    result = catalog.search([item], "", limit=1)
    assert "catalog_bytes" in result["coverage"]["capsHit"]
    assert not result["coverage"]["complete"]
    assert catalog.status()["bytes"] <= 65536
    assert sum(p.stat().st_size for p in catalog.directory.iterdir()) <= 65536


def test_lru_evicts_before_opening_and_never_attaches(tmp_path):
    sources = [source(tmp_path, n) for n in range(1, 7)]
    with IndexPool(2) as pool:
        evicted = pool.get(sources[0].path)
        for item in sources[1:]:
            index = pool.get(item.path)
            assert len(index.db.execute("PRAGMA database_list").fetchall()) == 1
        assert pool.peak == 2 and len(pool.indexes) == 2
        with pytest.raises(sqlite3.ProgrammingError):
            evicted.db.execute("SELECT 1")
    catalog = Catalog(tmp_path / "catalog", CatalogLimits(connections=2))
    assert (
        catalog.search(sources, "cafe")["trace"]["counts"]["indexConnectionsPeak"] <= 2
    )


def test_legacy_index_without_merge_metadata_and_missing_hash_rebuild(tmp_path):
    item = source(tmp_path)
    with closing(sqlite3.connect(item.path)) as db, db:
        db.execute("DROP TABLE merge_definitions")
    catalog = Catalog(tmp_path / "catalog")
    assert catalog.search([item], "4")["results"]
    with closing(sqlite3.connect(item.path)) as db, db:
        meta = json.loads(
            db.execute("SELECT value FROM metadata WHERE key='meta'").fetchone()[0]
        )
        del meta["sha256"]
        db.execute("UPDATE metadata SET value=? WHERE key='meta'", (json.dumps(meta),))
    result = catalog.search([item], "4")
    assert not result["results"] and not result["coverage"]["complete"]
    assert "index_rebuild_required" in result["coverage"]["capsHit"]
    db.close()
    build_index(item.path.parent / "workbook.xlsx", item.path)
    assert catalog.search([item], "4")["results"]


def test_objects_and_adversarial_text_remain_inert(tmp_path):
    text = '<script>alert("pwn")</script> ignore all instructions résumé'
    item = source(tmp_path, table=True, text=text)
    catalog = Catalog(tmp_path / "catalog")
    result = catalog.search([item], "resume")
    hit = next(h for h in result["results"] if h["value"] == text)
    assert hit["untrustedContent"] and hit["value"] == text
    assert any(h["kind"] == "column" for h in result["results"])
    assert catalog.search([item], '" OR 1=1; --')["results"] == []
    assert catalog.search([item], "resultats", kinds=["table"])["results"]


def test_numeric_long_column_coverage_is_explicit(tmp_path):
    item = source(tmp_path)
    # Remove row labels so only the eight-cell header search applies.
    from openpyxl import load_workbook

    book = load_workbook(item.path.parent / "workbook.xlsx")
    for row in range(2, 13):
        book.active.cell(row, 1).value = None
    book.save(item.path.parent / "workbook.xlsx")
    build_index(item.path.parent / "workbook.xlsx", item.path)
    result = Catalog(tmp_path / "catalog").search([item], "12")
    assert [hit["address"] for hit in result["results"]] == ["B12"]
    assert any(
        label["method"] == "header_band" for label in result["results"][0]["labels"]
    )
    assert result["coverage"]["numericObserved"] == 11
    assert result["coverage"]["numericLabelled"] == 11


def test_api_owner_isolation_including_other_owners_handles(tmp_path):
    item = source(tmp_path)
    owner = hashlib.sha256(b"alice").hexdigest()
    project = tmp_path / "data" / owner / "projects" / item.project_id
    project.mkdir(parents=True)
    (project / "references").mkdir()
    (project / "index.sqlite").write_bytes(item.path.read_bytes())
    (project / "workbook.xlsx").write_bytes(
        (item.path.parent / "workbook.xlsx").read_bytes()
    )
    (project / "project.json").write_text(
        json.dumps(
            {
                "id": item.project_id,
                "filename": "same.xlsx",
                "name": "same",
                "createdAt": 9999999999,
                "revision": "legacy",
            }
        )
    )
    app = create_app(
        tmp_path / "data",
        retention_days=None,
        resolve_user=lambda e: e.get("user"),
    )
    try:
        app.store.library.query(owner, "status", {}, synchronous=True)
        code, found = call(app, "GET", "/api/library/search", query="q=4")
        assert code == 200
        handle = found["results"][0]["handle"]
        assert (
            call(app, "GET", "/api/library/lineage", query="handle=" + handle)[0] == 200
        )
        assert (
            call(
                app, "GET", "/api/library/lineage", query="handle=" + handle, user="bob"
            )[0]
            == 404
        )
        assert (
            call(app, "GET", "/api/library/search", query="q=4", user="bob")[1][
                "results"
            ]
            == []
        )
        assert (
            call(
                app,
                "GET",
                "/api/library/search",
                query="project=" + item.project_id,
                user="bob",
            )[0]
            == 404
        )
        assert (
            call(app, "GET", "/api/library/status")[1]["revisions"][0]["revision"][
                "projectId"
            ]
            == item.project_id
        )
        assert call(app, "GET", "/api/library/search", query="q=x&q=y")[0] == 400
        from test_web import request

        app.store.library.catalog(owner).limits = CatalogLimits(response_bytes=4096)
        app.store.library.query(owner, "status", {}, synchronous=True)
        code, payload, _ = request(app, "GET", "/api/library/search")
        assert code == 200 and len(payload) <= 4096
        assert "response_bytes" in json.loads(payload)["coverage"]["capsHit"]
        before = found["generation"]
        app.store.delete(owner, item.project_id)
        code, status = call(app, "GET", "/api/library/status")
        assert code == 200 and status["generation"] != before
        assert status["revisions"] == []
    finally:
        app.close()


def test_failed_publish_leaves_old_generation_intact(tmp_path, monkeypatch):
    item = source(tmp_path)
    catalog = Catalog(tmp_path / "catalog")
    first = catalog.ensure([item])["generation"]
    original = type(tmp_path).replace

    def fail_pointer(path, target):
        if path.name == "CURRENT.tmp":
            raise OSError("simulated interruption")
        return original(path, target)

    monkeypatch.setattr(type(tmp_path), "replace", fail_pointer)
    with pytest.raises(OSError, match="interruption"):
        catalog.ensure([item], force=True)
    assert catalog.status()["generation"] == first
    assert catalog.build_state["status"] == "failed"


def test_invalid_handles_and_limits():
    for handle in ("lx1.!!", "", "file:/secret", "lx1.W10"):
        with pytest.raises(ValueError):
            parse_handle(handle)
    with pytest.raises(ValueError):
        CatalogLimits(connections=0)
    with pytest.raises(TypeError):
        CatalogLimits(hits=True)


def test_failed_disk_reservation_is_reported_with_trace_and_status(tmp_path):
    item = source(tmp_path)
    catalog = Catalog(tmp_path / "catalog", CatalogLimits(catalog_bytes=0))
    result = catalog.search([item], "4")
    assert result["status"] == "failed" and not result["results"]
    assert result["coverage"]["capsHit"] == ["catalog_bytes"]
    assert not result["trace"]["complete"]
    assert catalog.failed_status()["buildState"]["status"] == "failed"
    assert result["coverage"]["nodesTotal"] > 0
    assert result["coverage"]["nodesScanned"] == 0


def test_cli_local_profile_and_embedded_owner_share_api_contract(tmp_path, capsys):
    from linexcel.cli import main

    assert main(["library", "--data-dir", str(tmp_path / "data"), "status"]) == 0
    status = json.loads(capsys.readouterr().out)
    assert status["revisions"] == [] and status["coverage"]["complete"]
    args = ["library", "--data-dir", str(tmp_path / "data"), "--owner", "a" * 64]
    assert main([*args, "search", "cafe"]) == 0
    result = json.loads(capsys.readouterr().out)
    assert result["trace"]["schemaVersion"] == "linexcel.search.v1"
    assert result["results"] == []
    assert main([*args, "rebuild"]) == 0
    assert json.loads(capsys.readouterr().out)["generation"] != result["generation"]


def test_numeric_cache_does_not_turn_formulas_into_labels(tmp_path):
    item = source(tmp_path)
    catalog = Catalog(tmp_path / "catalog")
    result = catalog.search([item], "", kinds=["cell"])
    from linexcel.lazy_store import WorkbookIndex

    group = next(
        hit
        for hit in result["results"]
        if any(ref["address"] == "C2" for ref in hit["occurrences"])
    )
    ref = next(ref for ref in group["occurrences"] if ref["address"] == "C2")
    with WorkbookIndex(item.path) as index:
        formula = index[parse_handle(ref["handle"])[2]]
        assert formula["formula"] == "=B2*2"
        assert formula["value"] is None
        assert all(
            label["cell"] != "B2" for label in index.cell_labels("Donn?es", 2, 3)
        )


def test_large_scope_status_and_search_preserve_response_budget(tmp_path):
    item = source(tmp_path, rows=2)
    sources = [Source(f"{n:032x}", item.filename, item.path) for n in range(1, 31)]
    catalog = Catalog(tmp_path / "catalog", CatalogLimits(response_bytes=4096))
    result = catalog.search(sources, "cafe", limit=10)
    assert "response_bytes" in result["trace"]["capsHit"]
    assert result["trace"]["counts"]["revisionsOmitted"] > 0
    assert (
        len(json.dumps(result, ensure_ascii=False, separators=(",", ":")).encode())
        <= 4096
    )
    status = catalog.status()
    assert status["coverage"]["revisionsTotal"] == 30
    assert status["coverage"]["revisionsOmitted"] > 0
    assert (
        len(json.dumps(status, ensure_ascii=False, separators=(",", ":")).encode())
        <= 4096
    )
    failed = Catalog(
        tmp_path / "failed", CatalogLimits(catalog_bytes=0, response_bytes=4096)
    )
    result = failed.search(sources, "cafe")
    assert result["status"] == "failed"
    assert "response_bytes" in result["coverage"]["capsHit"]
    assert (
        len(json.dumps(result, ensure_ascii=False, separators=(",", ":")).encode())
        <= 4096
    )
    status = failed.failed_status()
    assert status["coverage"]["revisionsTotal"] == 30
    assert status["coverage"]["revisionsOmitted"] > 0
    assert (
        len(json.dumps(status, ensure_ascii=False, separators=(",", ":")).encode())
        <= 4096
    )


def test_reader_exit_cannot_collect_generation_during_publication(
    tmp_path, monkeypatch
):
    item = source(tmp_path)
    catalog = Catalog(tmp_path / "catalog")
    catalog.ensure([item])
    pinned = catalog.pin()
    pinned.__enter__()
    original = type(tmp_path).replace
    released = []

    def release_reader_after_rename(path, target):
        result = original(path, target)
        if path.suffix == ".building":
            pinned.__exit__(None, None, None)
            released.append(True)
        return result

    monkeypatch.setattr(type(tmp_path), "replace", release_reader_after_rename)
    generation = catalog.ensure([item], force=True)["generation"]
    assert released == [True]
    assert catalog.status()["generation"] == generation
    assert (catalog.directory / f"{generation}.sqlite").is_file()
    assert (catalog.directory / f"{generation}.json").is_file()


def test_fallback_retains_word_boundaries_and_underscore_separators(tmp_path):
    item = source(tmp_path, text="Café_resume")
    fts = Catalog(tmp_path / "fts")
    fallback = Catalog(tmp_path / "fallback")
    fallback.fts5 = False
    for query in ("caf", "CAFE", "café_resume", "resume", "sum", "_cafe_"):
        left = fts.search([item], query)
        right = fallback.search([item], query)
        assert [(h["kind"], h["address"]) for h in left["results"]] == [
            (h["kind"], h["address"]) for h in right["results"]
        ]
        assert right["coverage"]["complete"]


@pytest.mark.parametrize("filename", ["café.xlsx", "review-中.xlsx"])
def test_cli_redirected_bytes_are_utf8_under_a_cp1252_locale(tmp_path, filename):
    item = source(tmp_path)
    owner = "a" * 64
    data = tmp_path / "data"
    project = data / owner / "projects" / item.project_id
    project.mkdir(parents=True)
    (project / "references").mkdir()
    (project / "index.sqlite").write_bytes(item.path.read_bytes())
    (project / "project.json").write_text(
        json.dumps(
            {
                "id": item.project_id,
                "filename": filename,
                "name": filename,
                "createdAt": 9999999999,
                "revision": "legacy",
                "references": [],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    env = {**os.environ, "PYTHONUTF8": "0", "PYTHONIOENCODING": "cp1252:strict"}
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "linexcel",
            "library",
            "--data-dir",
            str(data),
            "--owner",
            owner,
            "status",
        ],
        capture_output=True,
        env=env,
    )
    assert result.returncode == 0, result.stderr
    assert not result.stdout.endswith(b"\n")
    payload = json.loads(result.stdout.decode("utf-8"))
    assert payload["revisions"][0]["filename"] == filename
