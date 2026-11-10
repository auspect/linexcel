"""A reader searches what the sheet shows: a displayed number or a cell address."""

import pytest
from openpyxl import Workbook

from linexcel.catalog import Catalog, Source
from linexcel.catalog_query import cell_address, displayed_numbers
from linexcel.lazy_store import build_index


def inside(value, ranges):
    return any(low <= value <= high for low, high in ranges)


@pytest.mark.parametrize(
    "query,found,missed",
    [
        ("6600", [6600.000000000001, 6600.4], [6601, 6599.4]),
        ("967.5", [967.538, 967.45], [967.56, 967.4]),
        ("967,5", [967.538], [9675]),
        ("1 455 900", [1455900.0], [1455901]),
        ("1 455 900,50", [1455900.5], [1455900.6]),
        ("1,455,900.5", [1455900.5], [1455.9]),
        ("1.455.900,5", [1455900.5], [1.4559]),
        ("1.455.900", [1455900.0], [1.4559]),
        # 1,234 is a French decimal and an English thousand: both are searched.
        ("1,234", [1.234, 1234], [12.34]),
        ("34.3%", [0.343, 34.3], [3.43]),
        ("-12.50", [-12.5, -12.504], [12.5]),
        ("1e6", [1000000.0], [1000000.4]),
    ],
)
def test_a_number_stands_for_the_saved_values_that_display_as_typed(
    query, found, missed
):
    ranges = displayed_numbers(query)
    assert all(inside(value, ranges) for value in found)
    assert not any(inside(value, ranges) for value in missed)


@pytest.mark.parametrize(
    "query", ["", "revenue", "Q4", "B5", "12 rue", "1,2,3", "nan", "inf", "%", "1.2.3"]
)
def test_text_is_not_read_as_a_number(query):
    assert displayed_numbers(query) == []


@pytest.mark.parametrize(
    "query,expected",
    [
        ("B5", (None, 5, 2, "B5")),
        ("$b$5", (None, 5, 2, "B5")),
        ("Synthèse!B5", ("Synthèse", 5, 2, "B5")),
        ("'Données 2026'!$B$2", ("Données 2026", 2, 2, "B2")),
        ("'O''Brien'!A1", ("O'Brien", 1, 1, "A1")),
        ("XFD1048576", (None, 1048576, 16384, "XFD1048576")),
    ],
)
def test_one_cell_address_is_recognized(query, expected):
    assert cell_address(query) == expected


@pytest.mark.parametrize(
    "query", ["B5:C6", "B", "5", "B0", "XFE1", "A1048577", "revenue", "!B5", "ABCD1"]
)
def test_other_text_is_not_an_address(query):
    assert cell_address(query) is None


@pytest.fixture
def report(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Synthèse"
    sheet["A5"] = "three_file_chain"
    sheet["B5"] = 6600.000000000001
    sheet["A6"] = "Exportations"
    sheet["B6"] = 967.538
    sheet["A7"] = "Q4"
    sheet["B7"] = 12
    other = book.create_sheet("Annexe")
    other["A5"] = "Autre"
    other["B5"] = 41
    book.save(tmp_path / "rapport.xlsx")
    build_index(
        tmp_path / "rapport.xlsx", tmp_path / "index.sqlite", filename="rapport.xlsx"
    )
    return Catalog(tmp_path / "catalog"), [
        Source("1" * 32, "rapport.xlsx", tmp_path / "index.sqlite", 1)
    ]


def cells(result):
    return [(hit["sheet"], hit["address"]) for hit in result["results"]]


def test_a_displayed_number_finds_its_saved_value(report):
    catalog, sources = report
    found = catalog.search(sources, "6600")
    assert cells(found) == [("Synthèse", "B5")]
    assert found["results"][0]["value"] == 6600.000000000001
    assert found["trace"]["numericMatch"]["ranges"] == [[6599.5, 6600.5]]
    assert cells(catalog.search(sources, "967.5")) == [("Synthèse", "B6")]
    assert cells(catalog.search(sources, "967,5")) == [("Synthèse", "B6")]
    # More decimals narrow the match: 967.54 displays 967.538, 967.60 does not.
    assert cells(catalog.search(sources, "967.54")) == [("Synthèse", "B6")]
    assert cells(catalog.search(sources, "967.60")) == []
    assert "numericMatch" not in catalog.search(sources, "Exportations")["trace"]


def test_an_address_without_text_match_shows_the_cells_at_that_address(report):
    catalog, sources = report
    bare = catalog.search(sources, "B5")
    assert cells(bare) == [("Synthèse", "B5"), ("Annexe", "B5")]
    assert bare["pagination"]["matched"] == 2 and bare["coverage"]["complete"]
    assert bare["trace"]["addressMatch"] == {
        "rule": "no text matched; cells at the queried address",
        "address": "B5",
        "sheet": None,
    }
    hit = bare["results"][0]
    assert hit["matchedBy"] == "address" and hit["value"] == 6600.000000000001
    assert catalog.lineage(sources, hit["handle"])["graph"]["nodes"]
    qualified = catalog.search(sources, "synthese!b5")
    assert cells(qualified) == [("Synthèse", "B5")]
    assert qualified["trace"]["addressMatch"]["sheet"] == "synthese"
    assert cells(catalog.search(sources, "Annexe!$B$5")) == [("Annexe", "B5")]
    # Paging and kind filters still apply.
    assert cells(catalog.search(sources, "B5", offset=1, limit=1)) == [("Annexe", "B5")]
    assert cells(catalog.search(sources, "B5", kinds=["sheet"])) == []
    assert cells(catalog.search(sources, "Z99")) == []


def test_text_that_looks_like_an_address_is_still_searched_as_text(report):
    catalog, sources = report
    found = catalog.search(sources, "Q4")
    # The label cell A7 holds the text "Q4"; cell Q4 of each sheet is empty.
    assert cells(found) == [("Synthèse", "A7")]
    assert "addressMatch" not in found["trace"]
