"""Stored XML formula spelling is evidence, independent from engine syntax."""

import hashlib
import io
import json
import zipfile
from xml.etree import ElementTree as ET

from openpyxl import Workbook

from linexcel import ExecutionPolicy, analyze
from linexcel.formula_source import enrich_graph_sources
from linexcel.lazy import evaluate_node
from linexcel.lazy_store import build_index, evidence_graph, get_node, graph_page


def package(setup):
    book = Workbook()
    setup(book.active)
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def rewrite_sheet(data, change):
    with zipfile.ZipFile(io.BytesIO(data)) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    sheet = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    change(sheet)
    parts["xl/worksheets/sheet1.xml"] = ET.tostring(sheet)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for name, content in parts.items():
            target.writestr(name, content)
    return output.getvalue()


def test_lookups_parentheses_survive_canonical_analysis_and_serialization():
    original = "=" + "(" * 40 + "1+2" + ")" * 40

    def setup(sheet):
        sheet.title = "Lookups"
        sheet["B22"] = original

    data = package(setup)
    result = analyze(data, execution=ExecutionPolicy(isolated=False))
    node = next(n for n in result.nodes if n.get("addr") == "B22")
    assert node["sourceFormula"] == original
    assert node["canonicalFormula"] == node["formula"] == "=1 + 2"
    assert node["value"] == 3
    provenance = node["formulaProvenance"]
    assert provenance["xmlText"] == original[1:]
    assert provenance["workbookSha256"] == hashlib.sha256(data).hexdigest()
    assert provenance["worksheetPart"] == "xl/worksheets/sheet1.xml"
    assert (provenance["sheet"], provenance["cell"]) == ("Lookups", "B22")
    assert provenance["canonicalMatchesSource"] is False
    assert node["decompositionFormulaBasis"] == "canonicalFormula"
    restored = json.loads(json.dumps(result.to_dict()))
    assert next(n for n in restored["nodes"] if n.get("addr") == "B22") == node


def shared_package():
    def setup(sheet):
        sheet["A1"] = 1
        sheet["A2"] = 2
        sheet["B1"] = "=((A1 + 2))"
        sheet["B2"] = "=((A2 + 2))"
        sheet["B3"] = "=A3+2"
        sheet["C1"] = True
        sheet["D1"] = "#VALUE!"
        sheet["D1"].data_type = "e"
        sheet["E1"] = "#VALUE!"
        sheet["E1"].data_type = "s"

    def shared(sheet):
        cells = {c.get("r"): c for c in sheet.findall("{*}sheetData/{*}row/{*}c")}
        master = cells["B1"].find("{*}f")
        master.attrib.update(t="shared", si="9", ref="B1:B2")
        for address, index in [("B2", "9"), ("B3", "99")]:
            formula = cells[address].find("{*}f")
            formula.attrib.update(t="shared", si=index)
            formula.text = None
        cells["B1"].find("{*}v").text = "3"
        cells["B2"].find("{*}v").text = "4"

    return rewrite_sheet(package(setup), shared)


def test_shared_formula_slave_records_translation_without_inventing_xml_source(
    tmp_path,
):
    data = shared_package()
    path = tmp_path / "source.xlsx"
    path.write_bytes(data)
    index = tmp_path / "source.sqlite"
    build_index(path, index)
    master = get_node(index, "Sheet!B1")
    slave = get_node(index, "Sheet!B2")
    missing = get_node(index, "Sheet!B3")
    assert master["sourceFormula"] == "=((A1 + 2))"
    assert master["formulaProvenance"]["xmlAttributes"] == {
        "t": "shared",
        "si": "9",
        "ref": "B1:B2",
    }
    assert slave["sourceFormula"] is None
    assert slave["formula"] == slave["canonicalFormula"] == "=((A2 + 2))"
    shared = slave["formulaProvenance"]["shared"]
    assert shared["role"] == "slave"
    assert shared["masterCell"] == "B1"
    assert shared["masterXmlText"] == "((A1 + 2))"
    assert shared["translatedFormula"] == slave["formula"]
    assert shared["translationStatus"] == "translated"
    assert slave["formulaProvenance"]["xmlText"] == ""
    assert missing["sourceFormula"] is None
    assert missing["formulaProvenance"]["status"] == "unavailable"
    assert (
        missing["formulaProvenance"]["shared"]["translationStatus"] == "missing_master"
    )
    assert graph_page(index, query="Sheet!B2")["nodes"][0] == slave
    assert (
        next(
            n
            for n in evidence_graph(index, "Sheet!B2")["nodes"]
            if n["id"] == "Sheet!B2"
        )["formulaProvenance"]
        == slave["formulaProvenance"]
    )
    result = evaluate_node(path, "Sheet!B2", index_path=index)
    assert result["value"] == 4
    assert slave["sourceValueProvenance"]["storageType"] == "n"
    assert slave["sourceValueProvenance"]["rawXmlValue"] == "4"
    assert get_node(index, "Sheet!C1")["sourceValueProvenance"]["storageType"] == "b"
    assert get_node(index, "Sheet!D1")["cachedValueKind"] == "error"
    assert get_node(index, "Sheet!D1")["sourceValueProvenance"]["storageType"] == "e"
    assert (
        get_node(index, "Sheet!E1")["sourceValueProvenance"]["storageType"]
        == "inlineStr"
    )


def test_retained_group_samples_have_their_own_xml_location_and_formula():
    data = shared_package()
    sample = {"addr": "B2", "value": 4}
    node = {
        "kind": "group",
        "sheet": "Sheet",
        "addr": "B1",
        "formula": "=A1+2",
        "samples": [sample],
    }
    warnings = []
    enrich_graph_sources(data, [node], warnings)
    assert node["formula"] == "=A1+2"
    assert node["sourceFormula"] == "=((A1 + 2))"
    assert sample["sourceFormula"] is None
    assert sample["canonicalFormula"] is None
    assert sample["formulaProvenance"]["cell"] == "B2"
    assert sample["formulaProvenance"]["shared"]["masterCell"] == "B1"
    assert sample["formulaProvenance"]["shared"]["translatedFormula"] == "=((A2 + 2))"
    assert not warnings


def test_missing_xml_evidence_is_explicit_and_never_rewrites_formula():
    node = {"kind": "cell", "sheet": "Sheet", "addr": "A1", "formula": "=1+2"}
    warnings = []
    enrich_graph_sources(b"invalid zip", [node], warnings)
    assert node["formula"] == node["canonicalFormula"] == "=1+2"
    assert node["sourceFormula"] is None
    assert node["formulaProvenance"]["status"] == "unavailable"
    assert warnings


def test_omitted_xml_cell_addresses_are_inferred_without_claiming_stored_address():
    def omit_addresses(sheet):
        for cell in sheet.findall("{*}sheetData/{*}row/{*}c"):
            cell.attrib.pop("r", None)

    data = rewrite_sheet(shared_package(), omit_addresses)
    node = {"kind": "cell", "sheet": "Sheet", "addr": "B2", "formula": "=A2+2"}
    warnings = []
    enrich_graph_sources(data, [node], warnings)
    assert node["sourceFormula"] is None
    assert node["formulaProvenance"]["cell"] == "B2"
    assert node["formulaProvenance"]["xmlCellAddress"] is None
    assert node["formulaProvenance"]["cellAddressOrigin"] == "inferred"
    assert node["formulaProvenance"]["shared"]["masterCell"] == "B1"
    assert node["formulaProvenance"]["shared"]["translatedFormula"] == "=((A2 + 2))"
    assert not warnings


def test_external_literal_keeps_the_referenced_file_xml_value_provenance(tmp_path):
    def input_sheet(sheet):
        sheet.title = "Data"
        sheet["A1"] = 41

    reference = package(input_sheet)
    refs = tmp_path / "refs"
    refs.mkdir()
    (refs / "Ref.xlsx").write_bytes(reference)
    main = package(
        lambda sheet: setattr(sheet["B1"], "value", "='[Ref.xlsx]Data'!A1+1")
    )
    path = tmp_path / "main.xlsx"
    path.write_bytes(main)
    index = tmp_path / "main.sqlite"
    build_index(path, index, refs)
    node = get_node(index, "external:Ref.xlsx:Data!A1")
    source = node["sourceValueProvenance"]
    assert source["workbookSha256"] == hashlib.sha256(reference).hexdigest()
    assert source["workbookSha256"] != hashlib.sha256(main).hexdigest()
    assert (source["sheet"], source["cell"]) == ("Data", "A1")
    assert source["storageType"] == "n"
    assert source["rawXmlValue"] == "41"
    assert source["role"] == "input"
    result = evaluate_node(path, "Sheet!B1", refs_dir=refs, index_path=index)
    assert result["value"] == 42
    step = next(s for s in result["steps"] if s["nodeId"] == node["id"])
    assert step["sourceValueProvenance"] == source
