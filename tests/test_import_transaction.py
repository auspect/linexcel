"""The import reads a worksheet inside a transaction (batch 2 ter).

Outside a transaction, every SQLite statement takes and releases the database
file lock: about 200 µs a statement on Windows, against about 1 µs inside one
(the same cost as critical review F24). Before index schema 5 each literal
cell was inserted on its own, and that insert kept a write transaction open.
Grid storage writes literal cells in batches, so between two periodic commits
each shared-string lookup ran on its own: the public DGT election table
(18 million cells) took 3,508 s to import instead of about 650 s.
"""

import io
import sqlite3
import zipfile

from openpyxl import Workbook

from linexcel.lazy_store import WorkbookIndex, build_index

ROWS, COLUMNS, DISTINCT = 12_500, 2, 3_000
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def string_index(row, column_index):
    return (row * COLUMNS + column_index) % DISTINCT


def shared_string_workbook(path):
    """A sheet of ROWS x COLUMNS text cells, all saved as shared strings."""
    book = Workbook()
    book.active.title = "Data"
    book.active["A1"] = "seed"
    stream = io.BytesIO()
    book.save(stream)
    letters = "ABCDEFGH"[:COLUMNS]
    cells = "".join(
        f'<row r="{row}">'
        + "".join(
            f'<c r="{letter}{row}" t="s"><v>{string_index(row, index)}</v></c>'
            for index, letter in enumerate(letters)
        )
        + "</row>"
        for row in range(1, ROWS + 1)
    )
    sheet = (
        f'<worksheet xmlns="{MAIN}"><dimension ref="A1:{letters[-1]}{ROWS}"/>'
        f"<sheetData>{cells}</sheetData></worksheet>"
    )
    strings = (
        f'<sst xmlns="{MAIN}" count="{ROWS * COLUMNS}" uniqueCount="{DISTINCT}">'
        + "".join(f"<si><t>text {index}</t></si>" for index in range(DISTINCT))
        + "</sst>"
    )
    source = zipfile.ZipFile(io.BytesIO(stream.getvalue()))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            text = source.read(info.filename).decode("utf-8")
            if info.filename == "xl/worksheets/sheet1.xml":
                text = sheet
            elif info.filename == "xl/_rels/workbook.xml.rels":
                text = text.replace(
                    "</Relationships>",
                    '<Relationship Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/sharedStrings" '
                    'Target="sharedStrings.xml" Id="rIdSst" /></Relationships>',
                )
            elif info.filename == "[Content_Types].xml":
                text = text.replace(
                    "</Types>",
                    '<Override PartName="/xl/sharedStrings.xml" ContentType="'
                    "application/vnd.openxmlformats-officedocument.spreadsheetml."
                    'sharedStrings+xml"/></Types>',
                )
            target.writestr(info, text)
        target.writestr("xl/sharedStrings.xml", strings)
    return path


def test_shared_string_lookups_run_inside_a_transaction(tmp_path, monkeypatch):
    workbook = shared_string_workbook(tmp_path / "strings.xlsx")
    lookups = []
    connect = sqlite3.connect

    class Watched(sqlite3.Connection):
        def execute(self, sql, *args):
            if sql.startswith("SELECT value FROM strings"):
                lookups.append(self.in_transaction)
            return super().execute(sql, *args)

    def watched_connect(target, *args, **kwargs):
        # build_index writes a temporary ".building" file, then renames it.
        if str(target).endswith(".building"):
            kwargs["factory"] = Watched
        return connect(target, *args, **kwargs)

    monkeypatch.setattr(sqlite3, "connect", watched_connect)
    index = tmp_path / "index-v5.sqlite"
    build_index(workbook, index)
    monkeypatch.setattr(sqlite3, "connect", connect)

    assert len(lookups) == ROWS * COLUMNS
    outside = lookups.count(False)
    assert outside == 0, f"{outside} of {len(lookups)} lookups outside a transaction"
    # The text itself is unchanged.
    with WorkbookIndex(index) as store:
        assert store["Data!A1"]["value"] == f"text {COLUMNS % DISTINCT}"
        last = ROWS * COLUMNS + COLUMNS - 1
        assert store[f"Data!B{ROWS}"]["value"] == f"text {last % DISTINCT}"
