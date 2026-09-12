"""Empty XML cells cannot steal a following formula or trigger backtracking."""

from linexcel.engine import _CELL_RE


def test_self_closing_cells_are_not_formula_containers():
    xml = b'<c r="A1" s="2"/><c r="B1" s="2" /><c r="C1"><f>1+2</f></c>'
    cells = list(_CELL_RE.finditer(xml))
    assert len(cells) == 1
    assert cells[0].groups() == (b"C", b"1", b"<f>1+2</f>")


def test_long_run_of_empty_cells_has_no_matches():
    xml = b"".join(f'<c r="A{row}" s="2"/>'.encode() for row in range(1, 20_001))
    assert list(_CELL_RE.finditer(xml)) == []


def test_explicit_empty_cells_and_formula_cells_keep_their_coordinates():
    xml = b'<c r="A1"></c><c r="B1" t="n"><f>A1+3</f><v>3</v></c>'
    assert [m.groups() for m in _CELL_RE.finditer(xml)] == [
        (b"A", b"1", b""),
        (b"B", b"1", b"<f>A1+3</f><v>3</v>"),
    ]
