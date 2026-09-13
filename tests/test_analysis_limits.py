"""Public ceilings are validated, observable, and never produce partial proofs."""

import io
import zipfile

import pytest
from openpyxl import Workbook

from linexcel import analyze
from linexcel.analyzer import analyze_workbook
from linexcel.external import read_workbook_values
from linexcel.loader import _detect_epoch_1904, load_cached_values
from linexcel.structure import inspect_workbook


def workbook(cells):
    wb = Workbook()
    wb.active.title = "S"
    for address, value in cells.items():
        wb.active[address] = value
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def replace_part(data, part, transform):
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source:
        with zipfile.ZipFile(output, "w") as target:
            for name in source.namelist():
                content = source.read(name)
                target.writestr(name, transform(content) if name == part else content)
    return output.getvalue()


@pytest.mark.parametrize("entry", [analyze, analyze_workbook, inspect_workbook])
@pytest.mark.parametrize(
    "name",
    [
        "max_cells_per_sheet",
        "max_nodes_per_sheet",
        "max_chain_depth",
        "max_dense_cells",
    ],
)
@pytest.mark.parametrize(
    "value, error",
    [(-1, ValueError), (True, TypeError), (1.5, TypeError), ("2", TypeError)],
)
def test_invalid_limits_fail_before_reading_the_workbook(entry, name, value, error):
    with pytest.raises(error, match=name):
        entry(b"not a workbook", **{name: value})


@pytest.mark.parametrize("dense", [0, 100])
@pytest.mark.parametrize("limit", [0, 1, 2, 3])
def test_cache_prefix_including_error_overlay_matches_both_readers(dense, limit):
    data = workbook({"A1": 11, "B1": "=1/0", "C1": 33})
    data = replace_part(
        data,
        "xl/worksheets/sheet1.xml",
        lambda xml: xml.replace(
            b'<c r="B1"><f>1/0</f><v /></c>',
            b'<c r="B1" t="e"><f>1/0</f><v>#DIV/0!</v></c>',
        ),
    )
    warnings = []
    cached = load_cached_values(
        data, warnings, max_cells_per_sheet=limit, max_dense_cells=dense
    )
    assert [cached.get("S", 1, col) for col in (1, 2, 3)] == [
        value if col <= limit else None
        for col, value in enumerate([11, "#DIV/0!", 33], 1)
    ]
    if limit < 3:
        assert warnings


def test_cell_limit_can_end_partway_through_a_row_and_is_reported():
    result = analyze(
        workbook({"A1": "=1+1", "B1": "=3*3", "C1": "=4^2"}), max_cells_per_sheet=2
    )
    assert result.stats["totalFormulas"] == 2
    assert result.graph["meta"]["analysisLimits"]["cellsPerSheet"] == 2
    assert any("partway" in warning for warning in result.warnings)


@pytest.mark.parametrize("dense", [0, 100])
def test_styled_empty_columns_use_the_same_cache_prefix_in_both_readers(dense):
    wb = Workbook()
    wb.active.title = "S"
    wb.active["A1"] = 1
    wb.active["A2"] = 2
    wb.active["C2"].number_format = "0.00"
    output = io.BytesIO()
    wb.save(output)
    cache = load_cached_values(
        output.getvalue(), max_cells_per_sheet=2, max_dense_cells=dense
    )
    assert cache.get("S", 1, 1) == 1
    assert cache.get("S", 2, 1) is None


def test_targeted_budget_counts_cells_instead_of_their_coordinates():
    data = workbook({"A999": "=2+3"})
    result = analyze(data, targets=["S!A999"], max_cells_per_sheet=1)
    assert result.stats["totalFormulas"] == 1
    assert result.nodes[0]["value"] == 5
    with pytest.raises(ValueError, match="complete traced lineage"):
        analyze(data, targets=["S!A999"], max_cells_per_sheet=0)


def test_node_limit_aggregates_excess_formulas_and_zero_is_honored():
    result = analyze(workbook({"A1": "=1+2", "B1": "=3*4"}), max_nodes_per_sheet=0)
    assert any(node["kind"] == "misc" for node in result.nodes)
    assert not any(node["kind"] in {"cell", "group"} for node in result.nodes)


def test_inspection_reports_effective_limits():
    info = inspect_workbook(
        workbook({"A1": 1}),
        max_cells_per_sheet=0,
        max_nodes_per_sheet=2,
        max_chain_depth=3,
        max_dense_cells=0,
    )
    assert info["ceilings"] == {"cellsPerSheet": 0, "nodesPerSheet": 2, "denseCells": 0}
    assert info["recoveryDepth"] == 3
    assert info["sheets"][0]["truncated"]
    assert info["densePathRefused"]


def test_external_dense_limit_is_checked_before_reader_allocation(tmp_path):
    path = tmp_path / "external.xlsx"
    path.write_bytes(workbook({"A1": 1}))
    with pytest.raises(ValueError, match="declares a used range"):
        read_workbook_values(path, max_dense_cells=0)


def test_named_external_resolution_honors_dense_limit(tmp_path):
    (tmp_path / "Ref.xlsx").write_bytes(workbook({"A1": 17}))
    result = analyze(
        workbook({"B1": "='[Ref.xlsx]S'!A1"}),
        refs_dir=tmp_path,
        max_dense_cells=0,
    )
    assert any("Ref.xlsx' could not be read" in warning for warning in result.warnings)
    assert result.stats["externalWorkbooksRead"] == 0


def test_declared_external_resolution_honors_dense_limit(tmp_path):
    from linexcel.external import ExternalBook, resolve_books

    (tmp_path / "Ref.xlsx").write_bytes(workbook({"A1": 17}))
    book = ExternalBook(key="1", target="Ref.xlsx", name="Ref.xlsx")
    warnings = []
    resolve_books({"1": book}, tmp_path, warnings, max_dense_cells=0)
    assert book.path is None
    assert any("could not be read" in warning for warning in warnings)


def test_external_read_never_returns_a_partial_sheet(tmp_path, monkeypatch):
    monkeypatch.setattr("linexcel.external.MAX_EXTERNAL_CELLS", 1)
    path = tmp_path / "Ref.xlsx"
    path.write_bytes(workbook({"A1": 17, "B1": 23}))
    with pytest.raises(ValueError, match="partial external values were not used"):
        read_workbook_values(path)


@pytest.mark.parametrize(
    "xml, expected",
    [
        (
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<workbookPr other="a" date1904 = "true" /></workbook>',
            True,
        ),
        (
            '<w:workbook xmlns:w="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<w:workbookPr date1904="1"/></w:workbook>',
            True,
        ),
        ('<workbook><ext><workbookPr date1904="1"/></ext></workbook>', False),
        (
            '<workbook xmlns:x="urn:extension"><workbookPr x:date1904="1"/></workbook>',
            False,
        ),
    ],
)
def test_epoch_xml_uses_direct_workbook_properties_only(xml, expected):
    data = replace_part(workbook({}), "xl/workbook.xml", lambda _: xml.encode())
    assert _detect_epoch_1904(data) is expected
