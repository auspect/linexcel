"""Catalog fairness, source-only header inference, and review regressions."""

import json
import sqlite3
import threading
import time

import pytest
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table

from linexcel.catalog import Catalog, CatalogLimits, Source, parse_handle
from linexcel.header_bands import HeaderBands
from linexcel.lazy_store import WorkbookIndex, _pack, _unpack, build_index
from linexcel.web import TaskStore, WebConfig


def indexed(tmp_path, book):
    path = tmp_path / "source.xlsx"
    book.save(path)
    index = tmp_path / "source.sqlite"
    build_index(path, index)
    return Source("1" * 32, path.name, index)


def test_every_sheet_and_structural_kind_survives_a_small_body_cap(tmp_path):
    book = Workbook()
    for i in range(3):
        sheet = book.active if i == 0 else book.create_sheet()
        sheet.title = f"Sheet{i}"
        sheet.append(["Region", "Amount"])
        for row in range(2, 202):
            sheet.append([f"Region {row}", row])
        sheet.add_table(Table(displayName=f"Table{i}", ref="A1:B201"))
    book.defined_names.add(DefinedName("LastSheetName", attr_text="Sheet2!$B$201"))
    item = indexed(tmp_path, book)
    catalog = Catalog(tmp_path / "catalog", CatalogLimits(entries_per_workbook=40))
    catalog.ensure([item])
    with catalog.pin() as (db, _):
        assert db.execute("SELECT count(*) FROM entries").fetchone()[0] <= 40
        for kind in ("sheet", "table", "column", "name"):
            assert db.execute(
                "SELECT count(*) FROM entries WHERE kind=?", (kind,)
            ).fetchone()[0]
        assert {
            r[0]
            for r in db.execute("SELECT DISTINCT sheet FROM entries WHERE kind='cell'")
        } == {f"Sheet{i}" for i in range(3)}
    for sheet in catalog.status()["coverage"]["sheets"]:
        if sheet["sheet"] is None:
            continue
        assert sheet["nodesScanned"] == sheet["nodesTotal"]
        assert sheet["byKind"]["text"]["cellsOmitted"] > 0
        assert sheet["byKind"]["numeric"]["cellsObserved"] == 200
        assert (
            "entries_per_sheet" in sheet["capsHit"]
            or "entries_per_workbook" in sheet["capsHit"]
        )


def test_fifty_thousand_repeats_are_one_entry_with_exact_bounded_occurrences(tmp_path):
    book = Workbook(write_only=True)
    sheet = book.create_sheet("Repeated")
    for row in range(50_000):
        sheet.append(["Café" if row % 2 else "CAFE"])
    item = indexed(tmp_path, book)
    catalog = Catalog(tmp_path / "catalog", CatalogLimits(text_occurrences=3))
    result = catalog.search([item], "cafe", kinds=["cell"])
    assert len(result["results"]) == 1
    hit = result["results"][0]
    assert hit["occurrenceCount"] == 50_000
    assert hit["occurrenceCountComplete"]
    assert len(hit["occurrences"]) == 3
    assert hit["occurrencesOmitted"] == 49_997
    assert {parse_handle(r["handle"])[0] for r in hit["occurrences"]} == {
        item.project_id
    }
    assert catalog.status()["coverage"]["entries"] == 2


def test_numeric_sampling_spans_each_sheet_and_groups_repeated_labels(tmp_path):
    book = Workbook()
    for i in range(2):
        sheet = book.active if i == 0 else book.create_sheet()
        sheet.title = f"Sheet{i}"
        sheet.append(["Amount"])
        for row in range(2, 102):
            sheet.append([row])
        sheet.add_table(Table(displayName=f"Table{i}", ref="A1:A101"))
    item = indexed(tmp_path, book)
    cat = Catalog(tmp_path / "cat", CatalogLimits(numeric_cells_per_workbook=10))
    cat.ensure([item])
    with cat.pin() as (db, _):
        rows = db.execute(
            "SELECT e.sheet,n.address FROM numeric_cells n "
            "JOIN entries e ON n.entry=e.id"
        ).fetchall()
        assert len(rows) == 10
        for sheet in ("Sheet0", "Sheet1"):
            addresses = [int(address[1:]) for s, address in rows if s == sheet]
            assert min(addresses) <= 21 and max(addresses) == 101
        assert (
            db.execute("SELECT count(*) FROM entries WHERE kind='cell'").fetchone()[0]
            == 2
        )
    assert "numeric_cells_per_workbook" in cat.status()["coverage"]["capsHit"]


def test_reallocated_budget_does_not_claim_full_counts_for_late_groups(tmp_path):
    book = Workbook()
    first = book.active
    first.title = "A"
    for row in range(1, 2001):
        first.append(["Repeat" if row in {100, 1537} else f"Unique {row}"])
    second = book.create_sheet("B")
    for _ in range(1024):
        second.append(["Other"])
    item = indexed(tmp_path, book)
    cat = Catalog(tmp_path / "cat", CatalogLimits(entries_per_workbook=50))
    hit = cat.search([item], "repeat", kinds=["cell"])["results"][0]
    # The first occurrence was omitted before B returned spare capacity. The
    # indexed count is a true lower bound, never the full source population.
    assert hit["address"] == "A1537"
    assert hit["occurrenceCount"] == 1
    assert not hit["occurrenceCountComplete"]


def test_numeric_literal_axis_needs_corroboration_and_keeps_exact_sources(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.append(["Code", "Name", "All ages", "0", "1", "2", "3"])
    for row in range(2, 15):
        sheet.append(
            [
                f"E{row}",
                f"Place {row}",
                1000 + row,
                2000 + row,
                3000 + row,
                4000 + row,
                5000 + row,
            ]
        )
    item = indexed(tmp_path, book)
    cat = Catalog(tmp_path / "catalog")
    hit = cat.search([item], "5014")["results"][0]
    labels = hit["labels"]
    assert {(label["cell"], label["text"], label["role"]) for label in labels} == {
        ("G1", "3", "column"),
        ("A14", "E14", "row"),
        ("B14", "Place 14", "row"),
    }
    with WorkbookIndex(item.path) as index:
        assert index.cell_labels("Sheet", 14, 7)  # Local row label, still conservative.
        assert not any(
            label["method"] == "header_band"
            for label in index.cell_labels("Sheet", 14, 7)
        )
    assert cat.status()["coverage"]["numericLabelled"] == 65


@pytest.mark.parametrize("case", ["blank", "section", "merged", "numeric", "cache"])
def test_header_bands_reject_false_positive_boundaries(tmp_path, case):
    book = Workbook()
    sheet = book.active
    sheet.append(["Region", "Amount"])
    for row in range(2, 7):
        sheet.append([f"Place {row}", row])
    if case == "blank":
        sheet.append([None, None])
    elif case == "section":
        sheet.append(["Next section", None])
    elif case == "merged":
        sheet.append(["Next section", None])
        sheet.merge_cells("A7:B7")
    elif case == "numeric":
        sheet["A1"] = 2025
        sheet["B1"] = 2026
    else:
        sheet["A1"] = '="Region"'
        sheet["B1"] = '="Amount"'
    for row in range(8, 14):
        sheet.cell(row, 2).value = row
    item = indexed(tmp_path, book)
    with WorkbookIndex(item.path) as index:
        bands = HeaderBands(index, "Sheet", [], CatalogLimits(), lambda: False)
        assert bands.labels(13, 2) == []
        if case in {"numeric", "cache"}:
            assert not bands.bands


def test_overlapping_merged_header_is_reported_ambiguous(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.append(["Region", "Amount"])
    sheet.merge_cells("A1:A2")
    sheet.append([None, 2])
    sheet.append(["Place", 3])
    item = indexed(tmp_path, book)
    with WorkbookIndex(item.path) as index:
        bands = HeaderBands(index, "Sheet", [], CatalogLimits(), lambda: False)
        assert not bands.bands
        assert bands.coverage["ambiguousCandidates"] == 1


def test_legacy_missing_merge_metadata_cannot_infer_unverified_headers(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.append(["Region", "Amount", None])
    sheet.merge_cells("B1:C1")
    for row in range(2, 15):
        sheet.append([f"Place {row}", row, row * 2])
    item = indexed(tmp_path, book)
    with WorkbookIndex(item.path, writable=True) as index:
        index.db.execute("DROP TABLE merge_definitions")
        index.db.commit()
        bands = HeaderBands(index, "Sheet", [], CatalogLimits(), lambda: False)
        assert not bands.bands
        assert bands.coverage["capsHit"] == ["header_merge_metadata_unavailable"]
        assert not bands.coverage["complete"]


def test_map_budget_never_extrapolates_and_is_stored_only_in_catalog(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.append(["Amount"])
    for row in range(2, 30):
        sheet.append([row])
    item = indexed(tmp_path, book)
    cat = Catalog(tmp_path / "catalog", CatalogLimits(header_scan_cells=5))
    assert not cat.search([item], "29")["results"]
    status = cat.status()
    assert "header_scan_cells" in status["coverage"]["capsHit"]
    with cat.pin() as (db, _):
        assert db.execute("SELECT count(*) FROM header_bands").fetchone()[0] == 1
        records = json.loads(db.execute("SELECT data FROM header_bands").fetchone()[0])
        assert records[0]["last_row"] <= 5
    with WorkbookIndex(item.path) as index:
        assert not index.db.execute(
            "SELECT 1 FROM sqlite_master WHERE name='header_bands'"
        ).fetchone()


@pytest.mark.parametrize(
    "budget",
    [
        "header_candidate_rows",
        "header_columns",
        "header_bands",
        "header_row_labels",
        "header_regions",
    ],
)
def test_each_header_budget_is_named_and_does_not_claim_complete(tmp_path, budget):
    book = Workbook()
    sheet = book.active
    sheet.append(["Code", "Name", "Amount"])
    for row in range(2, 15):
        sheet.append([f"E{row}", f"Place {row}", row])
    item = indexed(tmp_path, book)
    if budget == "header_regions":
        sheet.merge_cells("E1:F1")
        item = indexed(tmp_path, book)
    limits = CatalogLimits(**{budget: 0})
    result = Catalog(tmp_path / "cat", limits).search([item], "14")
    assert budget in result["coverage"]["capsHit"]
    assert not result["coverage"]["complete"]


def test_formula_cached_text_cannot_create_a_header_band(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.append(['="Region"', '="Amount"'])
    for row in range(2, 15):
        sheet.append([f"Place {row}", row])
    item = indexed(tmp_path, book)
    with WorkbookIndex(item.path, writable=True) as index:
        for col, value in ((1, "Region"), (2, "Amount")):
            seq, blob = index.db.execute(
                "SELECT seq,data FROM nodes WHERE row=1 AND col=?", (col,)
            ).fetchone()
            node = _unpack(blob)
            node.update(value=value, sourceDataType="str")
            index.db.execute("UPDATE nodes SET data=? WHERE seq=?", (_pack(node), seq))
        index.db.commit()
        bands = HeaderBands(index, "Sheet", [], CatalogLimits(), lambda: False)
        assert not bands.bands


def test_crash_staging_cleanup_preserves_current_and_pinned_generations(tmp_path):
    item = indexed(tmp_path, Workbook())
    cat = Catalog(tmp_path / "cat", CatalogLimits(catalog_bytes=256 * 1024))
    old = cat.ensure([item])["generation"]
    with cat.pin():
        (cat.directory / ("a" * 32 + ".building")).write_bytes(b"x" * 200_000)
        (cat.directory / "CURRENT.tmp").write_text("interrupted")
        (cat.directory / ("b" * 32 + ".sqlite")).write_bytes(b"x" * 200_000)
        new = cat.ensure([item], force=True)["generation"]
        assert new != old and (cat.directory / f"{old}.sqlite").exists()
        assert not list(cat.directory.glob("*.building"))
        assert not (cat.directory / ("b" * 32 + ".sqlite")).exists()
    assert cat.status()["generation"] == new


@pytest.mark.parametrize("artifact", ["sqlite", "json"])
def test_missing_published_artifact_rebuilds_without_leaking_pins(tmp_path, artifact):
    item = indexed(tmp_path, Workbook())
    cat = Catalog(tmp_path / "cat")
    old = cat.ensure([item])["generation"]
    (cat.directory / f"{old}.{artifact}").unlink()
    if artifact == "sqlite":
        with pytest.raises(sqlite3.OperationalError):
            with cat.pin():
                pass
        assert cat.pins == {}
    new = cat.ensure([item])["generation"]
    assert new != old and cat.status()["generation"] == new
    assert cat.pins == {}


def test_background_rebuild_does_not_hold_store_lock_or_block_old_search(
    tmp_path, monkeypatch
):
    item = indexed(tmp_path, Workbook())
    owner = "a" * 64
    store = TaskStore(WebConfig(tmp_path / "data", retention_days=None))
    started, release = threading.Event(), threading.Event()
    try:
        catalog = store.library.catalog(owner)
        catalog.ensure([item])
        monkeypatch.setattr(
            store.library,
            "sources",
            lambda requested: [item] if requested == owner else [],
        )
        original = catalog.ensure

        def slow(sources, **kwargs):
            catalog.build_state = {"status": "building", "phase": "test_gate"}
            started.set()
            assert release.wait(10)
            return original(sources, **kwargs)

        monkeypatch.setattr(catalog, "ensure", slow)
        with store.lock:
            store.library.refresh(owner)
        assert started.wait(5)
        before = time.monotonic()
        with store.lock:
            assert not store.closed
        result = store.library.query(owner, "search", {"q": [""]})
        status = store.library.query(owner, "status", {})
        assert time.monotonic() - before < 2
        assert result["stale"] and not result["coverage"]["complete"]
        assert result["generation"] == status["generation"]
        assert status["buildState"]["status"] == "building"
    finally:
        release.set()
        store.close()


def test_pending_first_search_keeps_trace_and_rejects_invalid_parameters(tmp_path):
    from linexcel.web import APIError

    store = TaskStore(WebConfig(tmp_path / "data", retention_days=None))
    owner = "a" * 64
    try:
        # A queued owner shares the existing bounded executor and has no
        # generation yet; no request may accidentally report complete absence.
        catalog = store.library.catalog(owner)
        store.library.pending[owner] = object()
        catalog.build_state = {"status": "queued"}
        result = store.library.query(owner, "search", {"q": ["amount"]})
        assert not result["coverage"]["complete"]
        assert result["trace"]["terms"] == ["amount"]
        assert result["trace"]["counts"]["returned"] == 0
        assert result["status"] == "building"
        with pytest.raises(APIError) as error:
            store.library.query(owner, "search", {"limit": ["0"]})
        assert error.value.status == 400
    finally:
        store.library.pending.clear()
        store.close()


def test_new_authorized_scope_returns_pending_while_old_generation_is_served(
    tmp_path, monkeypatch
):
    item = indexed(tmp_path, Workbook())
    new = Source("2" * 32, item.filename, item.path)
    owner = "a" * 64
    store = TaskStore(WebConfig(tmp_path / "data", retention_days=None))
    try:
        catalog = store.library.catalog(owner)
        catalog.ensure([item])
        monkeypatch.setattr(store.library, "sources", lambda _: [item, new])
        store.library.pending[owner] = object()
        catalog.build_state = {"status": "building"}
        result = store.library.query(owner, "search", {"project": [new.project_id]})
        assert result["status"] == "building" and result["results"] == []
        assert not result["coverage"]["complete"]
        assert result["trace"]["scope"]["project"] == new.project_id
    finally:
        store.library.pending.clear()
        store.close()
