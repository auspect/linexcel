"""What the node-list search (``graph_page`` with a query) finds.

Each case pins what the search returned on 3d2b2f9, before the search was
made faster, so that a faster search can be checked against it. The rules
are those of SQLite's ``LIKE`` on the ``nodes`` table, which the scan of the
literal cells (index schema 5) reproduces in Python: a substring of the id,
the formula or the text of the saved value, with ASCII letters folded only.
"""

from datetime import datetime

import pytest
from openpyxl import Workbook

from linexcel.lazy_store import build_index, graph_page


def search_workbook():
    book = Workbook()
    text = book.active
    text.title = "Texte"
    for value in [
        "Été",
        "été",
        "ÉTÉ",
        "Straße",
        "STRASSE",
        "strasse",
        "İstanbul",
        "istanbul",
        "ISTANBUL",
        "Euro EUR",
        "eur",
        "50%_off",
    ]:
        text.append([value])
    text["B1"] = '=LEN("Été")'
    plan = book.create_sheet("Plan 2")
    plan["A1"] = 12.5
    plan["A2"] = 0.1 + 0.2
    plan["A3"] = datetime(2024, 3, 15)
    plan["A4"] = True
    plan["A5"] = -7
    plan["A6"] = 1e20
    # Text in a cell formatted as a date, and text that reads as a boolean.
    plan["C1"] = "15/03/2024"
    plan["C1"].number_format = "dd/mm/yyyy"
    plan["C2"] = "TRUE"
    plan["B12"] = "x"
    plan["B120"] = "y"
    data = book.create_sheet("Données")
    for row in range(1, 31):
        data.cell(row, 1, f"ligne {row}")
    data["B1"] = "b"
    return book


#: (query, sheet filter, page size)
CASES = [
    # Case: ASCII letters fold, other letters do not.
    ("é", None, 200),
    ("É", None, 200),
    ("e", None, 200),
    ("E", None, 200),
    ("ß", None, 200),
    ("ss", None, 200),
    ("SS", None, 200),
    ("İ", None, 200),
    ("i", "Texte", 200),
    ("eur", None, 200),
    ("%_", None, 200),
    # Queries that look like an address, or end a sheet name.
    ("b12", None, 200),
    ("2!a1", None, 200),
    ("2'!a1", None, 200),
    ("es!", None, 200),
    ("es!b", None, 200),
    ("es'!b", None, 200),
    ("plan 2!", None, 200),
    ("texte!a1", None, 200),
    ("'plan", None, 200),
    # Numbers, dates and booleans, by their text.
    ("12.5", None, 200),
    ("0.3", None, 200),
    ("2024-03", None, 200),
    ("true", None, 200),
    ("-7", None, 200),
    ("e+20", None, 200),
    ("15/03", None, 200),
    ("03-15", None, 200),
    # Sheet filter.
    ("e", "Texte", 200),
    ("a1", "Données", 200),
    # A total past a full page.
    ("ligne", None, 5),
    ("ligne", "Données", 5),
    ("a", None, 5),
]

#: What the search returned on 3d2b2f9: (total, ids of the page).
EXPECTED = {
    ("é", None, 200): (
        34,
        [
            "Texte!A1",
            "Texte!B1",
            "Texte!A2",
            "'Données'!A1",
            "'Données'!B1",
            "'Données'!A2",
            "'Données'!A3",
            "'Données'!A4",
            "'Données'!A5",
            "'Données'!A6",
            "'Données'!A7",
            "'Données'!A8",
            "'Données'!A9",
            "'Données'!A10",
            "'Données'!A11",
            "'Données'!A12",
            "'Données'!A13",
            "'Données'!A14",
            "'Données'!A15",
            "'Données'!A16",
            "'Données'!A17",
            "'Données'!A18",
            "'Données'!A19",
            "'Données'!A20",
            "'Données'!A21",
            "'Données'!A22",
            "'Données'!A23",
            "'Données'!A24",
            "'Données'!A25",
            "'Données'!A26",
            "'Données'!A27",
            "'Données'!A28",
            "'Données'!A29",
            "'Données'!A30",
        ],
    ),
    ("É", None, 200): (3, ["Texte!A1", "Texte!B1", "Texte!A3"]),
    ("e", None, 200): (
        47,
        [
            "Texte!A1",
            "Texte!B1",
            "Texte!A2",
            "Texte!A3",
            "Texte!A4",
            "Texte!A5",
            "Texte!A6",
            "Texte!A7",
            "Texte!A8",
            "Texte!A9",
            "Texte!A10",
            "Texte!A11",
            "Texte!A12",
            "'Plan 2'!C2",
            "'Plan 2'!A4",
            "'Plan 2'!A6",
            "'Données'!A1",
            "'Données'!B1",
            "'Données'!A2",
            "'Données'!A3",
            "'Données'!A4",
            "'Données'!A5",
            "'Données'!A6",
            "'Données'!A7",
            "'Données'!A8",
            "'Données'!A9",
            "'Données'!A10",
            "'Données'!A11",
            "'Données'!A12",
            "'Données'!A13",
            "'Données'!A14",
            "'Données'!A15",
            "'Données'!A16",
            "'Données'!A17",
            "'Données'!A18",
            "'Données'!A19",
            "'Données'!A20",
            "'Données'!A21",
            "'Données'!A22",
            "'Données'!A23",
            "'Données'!A24",
            "'Données'!A25",
            "'Données'!A26",
            "'Données'!A27",
            "'Données'!A28",
            "'Données'!A29",
            "'Données'!A30",
        ],
    ),
    ("E", None, 200): (
        47,
        [
            "Texte!A1",
            "Texte!B1",
            "Texte!A2",
            "Texte!A3",
            "Texte!A4",
            "Texte!A5",
            "Texte!A6",
            "Texte!A7",
            "Texte!A8",
            "Texte!A9",
            "Texte!A10",
            "Texte!A11",
            "Texte!A12",
            "'Plan 2'!C2",
            "'Plan 2'!A4",
            "'Plan 2'!A6",
            "'Données'!A1",
            "'Données'!B1",
            "'Données'!A2",
            "'Données'!A3",
            "'Données'!A4",
            "'Données'!A5",
            "'Données'!A6",
            "'Données'!A7",
            "'Données'!A8",
            "'Données'!A9",
            "'Données'!A10",
            "'Données'!A11",
            "'Données'!A12",
            "'Données'!A13",
            "'Données'!A14",
            "'Données'!A15",
            "'Données'!A16",
            "'Données'!A17",
            "'Données'!A18",
            "'Données'!A19",
            "'Données'!A20",
            "'Données'!A21",
            "'Données'!A22",
            "'Données'!A23",
            "'Données'!A24",
            "'Données'!A25",
            "'Données'!A26",
            "'Données'!A27",
            "'Données'!A28",
            "'Données'!A29",
            "'Données'!A30",
        ],
    ),
    ("ß", None, 200): (1, ["Texte!A4"]),
    ("ss", None, 200): (2, ["Texte!A5", "Texte!A6"]),
    ("SS", None, 200): (2, ["Texte!A5", "Texte!A6"]),
    ("İ", None, 200): (1, ["Texte!A7"]),
    ("i", "Texte", 200): (2, ["Texte!A8", "Texte!A9"]),
    ("eur", None, 200): (2, ["Texte!A10", "Texte!A11"]),
    ("%_", None, 200): (1, ["Texte!A12"]),
    ("b12", None, 200): (2, ["'Plan 2'!B12", "'Plan 2'!B120"]),
    ("2!a1", None, 200): (0, []),
    ("2'!a1", None, 200): (1, ["'Plan 2'!A1"]),
    ("es!", None, 200): (0, []),
    ("es!b", None, 200): (0, []),
    ("es'!b", None, 200): (1, ["'Données'!B1"]),
    ("plan 2!", None, 200): (0, []),
    ("texte!a1", None, 200): (4, ["Texte!A1", "Texte!A10", "Texte!A11", "Texte!A12"]),
    ("'plan", None, 200): (
        10,
        [
            "'Plan 2'!A1",
            "'Plan 2'!C1",
            "'Plan 2'!A2",
            "'Plan 2'!C2",
            "'Plan 2'!A3",
            "'Plan 2'!A4",
            "'Plan 2'!A5",
            "'Plan 2'!A6",
            "'Plan 2'!B12",
            "'Plan 2'!B120",
        ],
    ),
    ("12.5", None, 200): (1, ["'Plan 2'!A1"]),
    ("0.3", None, 200): (1, ["'Plan 2'!A2"]),
    ("2024-03", None, 200): (1, ["'Plan 2'!A3"]),
    ("true", None, 200): (2, ["'Plan 2'!C2", "'Plan 2'!A4"]),
    ("-7", None, 200): (1, ["'Plan 2'!A5"]),
    ("e+20", None, 200): (1, ["'Plan 2'!A6"]),
    ("15/03", None, 200): (1, ["'Plan 2'!C1"]),
    ("03-15", None, 200): (1, ["'Plan 2'!A3"]),
    ("e", "Texte", 200): (
        13,
        [
            "Texte!A1",
            "Texte!B1",
            "Texte!A2",
            "Texte!A3",
            "Texte!A4",
            "Texte!A5",
            "Texte!A6",
            "Texte!A7",
            "Texte!A8",
            "Texte!A9",
            "Texte!A10",
            "Texte!A11",
            "Texte!A12",
        ],
    ),
    ("a1", "Données", 200): (
        11,
        [
            "'Données'!A1",
            "'Données'!A10",
            "'Données'!A11",
            "'Données'!A12",
            "'Données'!A13",
            "'Données'!A14",
            "'Données'!A15",
            "'Données'!A16",
            "'Données'!A17",
            "'Données'!A18",
            "'Données'!A19",
        ],
    ),
    ("ligne", None, 5): (
        30,
        [
            "'Données'!A1",
            "'Données'!A2",
            "'Données'!A3",
            "'Données'!A4",
            "'Données'!A5",
        ],
    ),
    ("ligne", "Données", 5): (
        30,
        [
            "'Données'!A1",
            "'Données'!A2",
            "'Données'!A3",
            "'Données'!A4",
            "'Données'!A5",
        ],
    ),
    ("a", None, 5): (52, ["Texte!A1", "Texte!A2", "Texte!A3", "Texte!A4", "Texte!A5"]),
}


@pytest.fixture(scope="module")
def index(tmp_path_factory):
    root = tmp_path_factory.mktemp("graph-page-search")
    book = root / "search.xlsx"
    search_workbook().save(book)
    path = root / "index.sqlite"
    build_index(book, path)
    return path


@pytest.mark.parametrize(("query", "sheet", "limit"), CASES)
def test_search_finds_what_it_found_before(index, query, sheet, limit):
    page = graph_page(index, query=query, sheet=sheet, limit=limit)
    total, ids = EXPECTED[(query, sheet, limit)]
    assert page["pagination"]["total"] == total
    assert [node["id"] for node in page["nodes"]] == ids


#: Texts and queries whose case folding differs between ASCII-only folding
#: and Unicode lower-casing (the Kelvin sign lowers to "k" in Python).
FOLDING = [
    "ÉtÉ",
    "Straße",
    "İstanbul",
    "K",
    "K",
    "ABC xyz [@`{]",
    "Ǆ",
    "ΣΑΣ",
    "",
]


@pytest.mark.parametrize("text", FOLDING)
@pytest.mark.parametrize("query", [*FOLDING[:-1], "k", "abc", "é", "ss"])
def test_the_scan_folds_case_as_sqlite_like_does(text, query):
    import sqlite3

    from linexcel.lazy_store import _filters, _like_search

    _, args = _filters(None, query, None)
    db = sqlite3.connect(":memory:")
    try:
        row = db.execute("SELECT ? LIKE ? ESCAPE '\\'", (text, args[0])).fetchone()
    finally:
        db.close()
    assert _like_search(query)(text) is bool(row[0])
