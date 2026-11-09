"""Regressions from the independent frozen M1b review's reproduced cases."""

import json
import shutil
import threading

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table

from linexcel.catalog import Catalog, CatalogLimits, Source, parse_handle
from linexcel.lazy_store import WorkbookIndex, _pack, build_index
from linexcel.web import APIError, TaskStore, WebConfig


def indexed(tmp_path, book):
    path = tmp_path / "source.xlsx"
    book.save(path)
    index = tmp_path / "source.sqlite"
    build_index(path, index)
    return Source("1" * 32, path.name, index)


@pytest.mark.parametrize("declared", [False, True])
def test_single_revenue_title_does_not_prove_a_column_label(tmp_path, declared):
    book = Workbook()
    for value in ("Revenue", 100, 120, 130):
        book.active.append([value])
    if declared:
        book.active.add_table(Table(displayName="ActualTable", ref="A1:A4"))
    item = indexed(tmp_path, book)
    cat = Catalog(tmp_path / "catalog")
    for value in (100, 120, 130):
        hits = cat.search([item], str(value))["results"]
        assert bool(hits) is declared
        if declared:
            assert {label["method"] for label in hits[0]["labels"]} == {"table_column"}


@pytest.mark.parametrize("declared", [False, True])
def test_comment_is_context_not_the_row_key_and_all_bands_are_heuristic(
    tmp_path, declared
):
    book = Workbook()
    book.active.append(["ID", "Comment", "Amount"])
    book.active.append(["A1", "estimated", 10])
    if declared:
        book.active.add_table(Table(displayName="ActualTable", ref="A1:C2"))
    item = indexed(tmp_path, book)
    cat = Catalog(tmp_path / "catalog")
    labels = cat.search([item], "10")["results"][0]["labels"]
    if declared:
        assert not any(label["role"] == "row" for label in labels)
        assert (
            next(label for label in labels if label["text"] == "estimated")["role"]
            == "row_context"
        )
        return
    assert {(label["text"], label["role"]) for label in labels} == {
        ("Amount", "column"),
        ("A1", "row"),
        ("estimated", "row_context"),
    }
    assert all(label["certainty"] == "heuristic" for label in labels)
    assert all(label.get("ambiguous") is not False for label in labels)
    assert all(label["method"] == "header_band" for label in labels)
    assert {(label["field"], label["fieldCell"]) for label in labels} == {
        ("Amount", "C1"),
        ("ID", "A1"),
        ("Comment", "B1"),
    }


def test_two_owners_building_cannot_starve_import_executor(tmp_path, monkeypatch):
    store = TaskStore(WebConfig(tmp_path / "data", retention_days=None, max_workers=2))
    started, release = threading.Event(), threading.Event()
    owners = ["a" * 64, "b" * 64]
    builds = []
    try:
        monkeypatch.setattr(store.library, "sources", lambda _: [])
        for owner in owners:

            def slow(sources, owner=owner):
                builds.append(owner)
                started.set()
                assert release.wait(10)

            monkeypatch.setattr(store.library.catalog(owner), "ensure", slow)
            store.library.refresh(owner)
        assert started.wait(5)
        assert (
            store.executor.submit(lambda: "import slot").result(timeout=2)
            == "import slot"
        )
        with store.lock:
            store.closed = True
        release.set()
        for future in list(store.library.pending.values()):
            future.result(timeout=5)
        assert builds == [
            owners[0]
        ]  # Queued owner exits before source enumeration/build.
    finally:
        release.set()
        store.close()


def test_closed_reconciliation_does_not_even_enumerate_sources(tmp_path, monkeypatch):
    store = TaskStore(WebConfig(tmp_path / "data", retention_days=None))
    try:
        store.closed = True
        monkeypatch.setattr(
            store.library, "mutation", lambda _: pytest.fail("closed gate")
        )
        monkeypatch.setattr(
            store.library, "sources", lambda _: pytest.fail("closed enumeration")
        )
        store.library._reconcile("a" * 64)
        assert not store.library.catalogs
    finally:
        store.close()


@pytest.mark.parametrize("failure", [OSError, ValueError, RuntimeError])
def test_snapshot_exception_is_a_terminal_failed_status(tmp_path, monkeypatch, failure):
    store = TaskStore(WebConfig(tmp_path / "data", retention_days=None))
    owner = "a" * 64
    try:

        def broken(_):
            raise failure("private source path must not be exposed")

        monkeypatch.setattr(store.library, "sources", broken)
        with store.lock:
            store.library.refresh(owner)
            future = store.library.pending[owner]
        future.result(timeout=5)
        assert owner not in store.library.pending
        result = store.library.query(owner, "status", {})
        assert result["buildState"]["status"] == "failed"
        assert result["buildState"]["phase"] == "sources"
        assert not result["coverage"]["complete"]
        assert "catalog_source_snapshot_failed" in result["coverage"]["capsHit"]
        assert "private source" not in json.dumps(result)
    finally:
        store.close()


def repeated(tmp_path, **budgets):
    book = Workbook()
    for _ in range(35):
        book.active.append(["Repeated"])
    item = indexed(tmp_path, book)
    # Obtain the group under a roomy search envelope, then apply tighter read
    # budgets to the same immutable generation for occurrence-page checks.
    cat = Catalog(
        tmp_path / "catalog", CatalogLimits(**{**budgets, "response_bytes": 1048576})
    )
    hit = cat.search([item], "Repeated", kinds=["cell"])["results"][0]
    cat = Catalog(tmp_path / "catalog", CatalogLimits(**budgets))
    return cat, item, hit


def test_group_pages_reach_every_omitted_occurrence_and_lineage_handle(tmp_path):
    cat, item, hit = repeated(tmp_path, text_occurrences=2, scan_entries=7)
    assert hit["occurrencesOmitted"] == 33
    group = hit["occurrenceGroup"]
    after, found = 0, []
    while True:
        page = cat.occurrences(
            [item],
            generation=group["generation"],
            group=group["id"],
            after=after,
            limit=3,
        )
        found.extend(page["results"])
        assert page["coverage"]["cellsScanned"] <= 7
        assert len(page["results"]) <= 3
        next_after = page["pagination"]["nextAfter"]
        if next_after is None:
            assert page["coverage"]["complete"]
            break
        assert next_after > after
        after = next_after
    assert {r["address"] for r in found} == {f"A{r}" for r in range(1, 36)}
    assert len(found) == 35
    last = next(r for r in found if r["address"] == "A35")
    assert cat.lineage([item], last["handle"])["graph"]["nodes"]


def test_occurrence_pages_reread_authoritative_values_and_reauthorize(tmp_path):
    cat, item, hit = repeated(tmp_path)
    args = dict(
        generation=hit["occurrenceGroup"]["generation"],
        group=hit["occurrenceGroup"]["id"],
    )
    _, _, nid = parse_handle(hit["handle"])
    with WorkbookIndex(item.path, writable=True) as index:
        node = index[nid]
        node["value"] = "Changed"
        index.db.execute("UPDATE nodes SET data=? WHERE id=?", (_pack(node), nid))
        index.db.commit()
    page = cat.occurrences([item], **args)
    assert len(page["results"]) == 34
    assert "A1" not in {r["address"] for r in page["results"]}
    with pytest.raises(KeyError):
        cat.occurrences([], **args)
    with pytest.raises(KeyError):
        cat.occurrences([Source("2" * 32, item.filename, item.path)], **args)
    with WorkbookIndex(item.path, writable=True) as index:
        meta = {**index.meta, "sha256": "f" * 64}
        index.db.execute(
            "UPDATE metadata SET value=? WHERE key='meta'", (json.dumps(meta),)
        )
        index.db.commit()
    with pytest.raises(KeyError):
        cat.occurrences([item], **args)


def test_occurrence_pages_are_generation_scoped_and_validate_budgets(tmp_path):
    cat, item, hit = repeated(tmp_path)
    args = dict(
        generation=hit["occurrenceGroup"]["generation"],
        group=hit["occurrenceGroup"]["id"],
    )
    for override in (
        {"group": 0},
        {"after": -1},
        {"limit": 501},
        {"generation": "../secret"},
        {"after": True},
    ):
        with pytest.raises(ValueError):
            cat.occurrences([item], **{**args, **override})
    cat.ensure([item], force=True)
    with pytest.raises(ValueError, match="generation changed"):
        cat.occurrences([item], **args)


def test_occurrence_api_is_owner_scoped_and_strict(tmp_path, monkeypatch):
    cat, item, hit = repeated(tmp_path)
    owner = "a" * 64
    store = TaskStore(WebConfig(tmp_path / "data", retention_days=None))
    try:
        store.library.catalogs[owner] = cat
        monkeypatch.setattr(
            store.library,
            "sources",
            lambda requested, search_scope=None: [item] if requested == owner else [],
        )
        params = {
            "generation": [hit["occurrenceGroup"]["generation"]],
            "group": [str(hit["occurrenceGroup"]["id"])],
            "limit": ["10"],
        }
        page = store.library.query(owner, "occurrences", params)
        assert len(page["results"]) == 10
        for supplied in ({**params, "owner": [owner]}, {**params, "group": ["1", "2"]}):
            with pytest.raises(APIError) as error:
                store.library.query(owner, "occurrences", supplied)
            assert error.value.status == 400
        # Reuse the exact generation and group under an owner with no sources.
        store.library.catalogs["b" * 64] = cat
        with pytest.raises(APIError) as error:
            store.library.query("b" * 64, "occurrences", params)
        assert error.value.status == 404
    finally:
        store.close()


def test_occurrence_response_and_scan_budgets_resume_without_losing_matches(tmp_path):
    cat, item, hit = repeated(tmp_path, response_bytes=4096, scan_entries=5)
    args = dict(
        generation=hit["occurrenceGroup"]["generation"],
        group=hit["occurrenceGroup"]["id"],
    )
    found, after = [], 0
    while True:
        page = cat.occurrences([item], **args, after=after, limit=500)
        assert (
            len(json.dumps(page, ensure_ascii=False, separators=(",", ":")).encode())
            <= 4096
        )
        assert page["coverage"]["cellsScanned"] <= 5
        found.extend(r["address"] for r in page["results"])
        next_after = page["pagination"]["nextAfter"]
        if next_after is None:
            break
        assert "scan_entries" in page["coverage"]["capsHit"]
        assert next_after > after
        after = next_after
    assert len(found) == len(set(found)) == 35


def test_occurrence_response_cap_resumes_at_last_returned_source_cell(tmp_path):
    cat, item, hit = repeated(tmp_path, response_bytes=4096)
    args = dict(
        generation=hit["occurrenceGroup"]["generation"],
        group=hit["occurrenceGroup"]["id"],
    )
    page = cat.occurrences([item], **args, limit=500)
    assert "response_bytes" in page["coverage"]["capsHit"]
    assert 0 < len(page["results"]) < 35
    next_page = cat.occurrences(
        [item], **args, after=page["pagination"]["nextAfter"], limit=500
    )
    assert set(r["address"] for r in page["results"]).isdisjoint(
        r["address"] for r in next_page["results"]
    )


def test_occurrence_expansion_recomputes_label_groups_beyond_numeric_sample(tmp_path):
    book = Workbook()
    book.active.append(["ID", "Comment", "Amount"])
    for row in range(2, 32):
        book.active.append(["A1", "estimated", row])
    item = indexed(tmp_path, book)
    cat = Catalog(tmp_path / "catalog", CatalogLimits(numeric_cells_per_workbook=12))
    hits = cat.search([item], "Amount estimated", kinds=["cell"])["results"]
    hit = hits[0]
    assert hit["occurrenceCount"] < 30
    assert not hit["occurrenceCountComplete"]
    page = cat.occurrences(
        [item],
        generation=hit["occurrenceGroup"]["generation"],
        group=hit["occurrenceGroup"]["id"],
        limit=100,
    )
    assert {ref["address"] for ref in page["results"]} == {
        f"C{r}" for r in range(2, 32)
    }
    assert page["coverage"]["complete"]


def test_occurrence_identity_uses_full_text_instead_of_truncated_search_text(tmp_path):
    book = Workbook()
    for value in ("same " * 1000 + "first", "same " * 1000 + "second"):
        book.active.append([value])
    item = indexed(tmp_path, book)
    cat = Catalog(tmp_path / "catalog")
    hits = cat.search([item], "same", kinds=["cell"])["results"]
    assert len(hits) == 2
    pages = [
        cat.occurrences(
            [item],
            generation=hit["occurrenceGroup"]["generation"],
            group=hit["occurrenceGroup"]["id"],
        )
        for hit in hits
    ]
    assert [{r["address"] for r in page["results"]} for page in pages] == [
        {"A1"},
        {"A2"},
    ]


def test_reconciliation_failure_before_ensure_sets_failed_state(tmp_path, monkeypatch):
    store = TaskStore(WebConfig(tmp_path / "data", retention_days=None))
    owner = "a" * 64
    try:
        cat = store.library.catalog(owner)

        def broken(*args):
            raise RuntimeError("manifest snapshot failed")

        monkeypatch.setattr(cat, "_manifest", broken)
        with store.lock:
            store.library.refresh(owner)
            future = store.library.pending[owner]
        future.result(timeout=5)
        status = store.library.query(owner, "status", {})
        assert status["buildState"]["status"] == "failed"
        assert not status["coverage"]["complete"]
    finally:
        store.close()


def test_occurrence_cli_reaches_omitted_handles(tmp_path, capsys):
    from linexcel.cli import main

    _, item, _ = repeated(tmp_path)
    owner = "a" * 64
    data = tmp_path / "workspace"
    project = data / owner / "projects" / item.project_id
    project.mkdir(parents=True)
    shutil.copyfile(item.path, project / "index.sqlite")
    (project / "project.json").write_text(
        json.dumps(
            {
                "id": item.project_id,
                "filename": item.filename,
                "name": "Repeated",
                "createdAt": 9999999999,
                "revision": "legacy",
                "references": [],
            }
        ),
        encoding="utf-8",
    )
    args = ["library", "--data-dir", str(data), "--owner", owner]
    assert main([*args, "search", "Repeated", "--kinds", "cell"]) == 0
    hit = json.loads(capsys.readouterr().out)["results"][0]
    group = hit["occurrenceGroup"]
    assert (
        main(
            [
                *args,
                "occurrences",
                str(group["id"]),
                "--generation",
                group["generation"],
                "--limit",
                "500",
            ]
        )
        == 0
    )
    page = json.loads(capsys.readouterr().out)
    assert len(page["results"]) == 35
    assert any(r["address"] == "A35" for r in page["results"])


@pytest.mark.parametrize("budget", ["query_ms", "scan_entries"])
def test_occurrence_zero_budget_fails_without_a_nonadvancing_cursor(tmp_path, budget):
    cat, item, hit = repeated(tmp_path)
    cat = Catalog(cat.directory, CatalogLimits(**{budget: 0}))
    with pytest.raises(ValueError, match="cannot advance"):
        cat.occurrences(
            [item],
            generation=hit["occurrenceGroup"]["generation"],
            group=hit["occurrenceGroup"]["id"],
        )
