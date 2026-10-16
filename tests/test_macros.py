"""Static add-in reading must preserve evidence and never invoke Excel."""

import hashlib
import json
from pathlib import Path

import pytest

import linexcel
from linexcel import ExecutionPolicy
from linexcel.cli import main
from linexcel.lazy_store import build_index, export_structure, get_node, structure_page
from linexcel.macros import OLE_MAGIC, valid_excel_header


@pytest.fixture
def legacy_source(monkeypatch):
    import oletools.olevba

    code = (
        'Attribute VB_Name = "Module1"\n'
        'Sub Refresh()\n  Helper\n  Worksheets("Summary").Range("A1") = 2\nEnd Sub\n'
        'Sub Helper()\n  x = Range("B3")\nEnd Sub\n'
        "Property Get Rate() As Double\n  Rate = 1\nEnd Property\n"
        "Property Let Rate(ByVal v As Double)\n  x = v\nEnd Property\n"
    )
    closed = []

    class Parser:
        def __init__(self, filename, data):
            assert filename.endswith(".xla")
            assert data.startswith(OLE_MAGIC)

        def detect_vba_macros(self):
            return True

        def extract_macros(self):
            yield "tool.xla", "VBA/Module1", "Module1.bas", code

        def close(self):
            closed.append(True)

    monkeypatch.setattr(oletools.olevba, "VBA_Parser", Parser)
    return OLE_MAGIC + b"static test project", code, closed


def test_real_xlam_source_has_binary_and_decoded_text_fingerprints(monkeypatch):
    def forbidden(*args, **kwargs):
        raise AssertionError("Reading VBA must not launch Office or execute code")

    monkeypatch.setattr("subprocess.run", forbidden)
    data = Path("tests/fixtures/macros.xlsm").read_bytes()
    source = linexcel.read_macros(data, filename="tools.xlam")
    assert source["execution"] == "not_executed"
    assert source["sha256"] == hashlib.sha256(data).hexdigest()
    assert source["procedures"]
    assert not source["warnings"]
    for module, code in source["modules"].items():
        origin = source["moduleProvenance"][module]
        assert origin["moduleSha256"] == hashlib.sha256(code.encode()).hexdigest()
        assert origin["sourceFile"] == "tools.xlam"


def test_xla_source_graph_keeps_properties_and_explicit_coverage(legacy_source):
    data, code, closed = legacy_source
    source = linexcel.read_macros(data, filename="tool.xla")
    assert source["modules"] == {"Module1": code}
    assert len(source["procedures"]) == 4
    result = linexcel.analyze(
        data, filename="tool.xla", execution=ExecutionPolicy(isolated=False)
    )
    nodes = {n["id"]: n for n in result.nodes}
    assert len([n for n in nodes.values() if n["kind"] == "vba"]) == 4
    assert result.graph["meta"]["macroOnly"]
    assert result.graph["meta"]["analysisCoverage"]["worksheetCells"] == "not_analyzed"
    assert result.engine is None
    assert result.workbook_context["sheets"] == []
    assert any(
        e["kind"] == "call" and e["target"] == "vp:Module1.Helper" for e in result.edges
    )
    assert {e["kind"] for e in result.edges} == {"call", "vba-read", "vba-write"}
    assert all(
        e["approx"] and e["scope"] == "static_vba_heuristic" for e in result.edges
    )
    assert all(n.get("value") is None for n in nodes.values())
    assert len(closed) == 2


def test_legacy_index_preserves_static_edge_kinds_and_source(legacy_source, tmp_path):
    data, _, _ = legacy_source
    path = tmp_path / "workbook.xlsx"
    path.write_bytes(data)
    index = tmp_path / "index.sqlite"
    graph = build_index(path, index, filename="tool.xla")
    assert graph["meta"]["macroOnly"]
    assert graph["meta"]["filename"] == "tool.xla"
    assert graph["meta"]["vba"]["procedures"] == 4
    node = get_node(index, "vp:Module1.Refresh")
    assert node["vbaProvenance"]["sourceFile"] == "tool.xla"
    assert node["vbaProvenance"]["sourceSha256"] == hashlib.sha256(data).hexdigest()
    assert all(
        e["approx"] and e["scope"] == "static_vba_heuristic"
        for e in export_structure(index)["edges"]
    )
    assert all(
        e["approx"] and e["scope"] == "static_vba_heuristic"
        for e in structure_page(index, part="edges")["edges"]
    )
    assert {e["kind"] for e in structure_page(index, part="edges")["edges"]} == {
        "call",
        "vba-read",
        "vba-write",
    }


def test_header_does_not_admit_legacy_xls_as_a_workbook():
    assert valid_excel_header(OLE_MAGIC, "tool.xla")
    assert not valid_excel_header(OLE_MAGIC, "report.xlsm")
    assert not valid_excel_header(OLE_MAGIC, "report.xls")
    assert valid_excel_header(b"PK\x03\x04", "tool.xlam")
    assert not valid_excel_header(b"PK\x03\x04", "tool.xla")
    with pytest.raises(ValueError):
        linexcel.read_macros(OLE_MAGIC, filename="report.xls")


def test_scanner_does_not_invent_string_reads_or_conditional_writes():
    from linexcel.vba import analyze_vba

    proc = analyze_vba(
        {
            "M": 'Sub S()\n MsgBox "See [A1]"\n '
            'If Range("B1").Value = 2 Then\n End If\n '
            'If True Then Range("C1").Value = 2\n '
            'x = Range("D1").Value = 2\n '
            'ActiveWorkbook.Worksheets("Summary").Range("E1") = 3\nEnd Sub'
        }
    )[0]
    assert [(r.ref, r.access) for r in proc.refs] == [
        ("B1", "read"),
        ("C1", "write"),
        ("D1", "read"),
        ("E1", "write"),
    ]
    assert proc.refs[-1].workbook == "ActiveWorkbook"


def test_actual_xlam_content_type_uses_static_addin_api():
    import io
    import zipfile

    data = Path("tests/fixtures/macros.xlsm").read_bytes()
    result_file = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(data)) as source,
        zipfile.ZipFile(result_file, "w") as out,
    ):
        for item in source.infolist():
            payload = source.read(item.filename)
            if item.filename == "[Content_Types].xml":
                payload = payload.replace(
                    b"application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                    b"application/vnd.ms-excel.addin.macroEnabled.main+xml",
                )
            out.writestr(item, payload)
    result = linexcel.analyze(
        result_file.getvalue(),
        filename="tools.xlam",
        execution=ExecutionPolicy(isolated=False),
    )
    assert result.graph["meta"]["macroOnly"]
    assert result.stats["vbaProcs"] > 0
    assert result.graph["meta"]["analysisCoverage"]["worksheetCells"] == "not_analyzed"


def test_macro_stream_collision_keeps_all_origins_and_warning(monkeypatch):
    import oletools.olevba

    from linexcel.vba import extract_vba_modules

    class Parser:
        def __init__(self, *args, **kwargs):
            pass

        def detect_vba_macros(self):
            return True

        def extract_macros(self):
            yield "x", "Project1/VBA/M", "M.bas", "Sub A()\nEnd Sub"
            yield "x", "Project2/VBA/M", "M.bas", "Sub B()\nEnd Sub"

        def close(self):
            pass

    monkeypatch.setattr(oletools.olevba, "VBA_Parser", Parser)
    warnings, origins = [], {}
    modules = extract_vba_modules(b"PK test", "tool.xlam", warnings, provenance=origins)
    assert "Sub A()" in modules["M"] and "Sub B()" in modules["M"]
    assert len(origins["M"]["streamOrigins"]) == 2
    assert origins["M"]["streamPath"] is None
    assert origins["M"]["moduleMerged"] and warnings


def test_macros_cli_writes_complete_source_report(legacy_source, tmp_path):
    data, code, _ = legacy_source
    path, report = tmp_path / "tool.xla", tmp_path / "source.json"
    path.write_bytes(data)
    assert main(["macros", str(path), "-o", str(report)]) == 0
    assert json.loads(report.read_text(encoding="utf-8"))["modules"]["Module1"] == code


def test_static_xla_rejects_cell_target(legacy_source):
    with pytest.raises(ValueError, match="Cell targets"):
        linexcel.analyze(
            legacy_source[0],
            filename="tool.xla",
            targets=["Sheet1!A1"],
            execution=ExecutionPolicy(isolated=False),
        )
