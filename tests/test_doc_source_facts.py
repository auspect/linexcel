"""Source facts that generated prose previously misrepresented in local audits."""

import io
import json
import subprocess
import sys
import zipfile

import pytest

from linexcel.aidoc import (
    MAX_WORKBOOK_DOSSIER_CHARS,
    _fit_workbook_dossier,
    build_dossier,
    build_workbook_dossier,
)
from linexcel.doc_evidence import (
    formula_evidence,
    read_defined_name_evidence,
    relevant_names,
)


def package(names, *, sheets='<sheet name="Data"/><sheet name="Other"/>'):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w") as z:
        z.writestr(
            "xl/workbook.xml",
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"<sheets>{sheets}</sheets><definedNames>{names}</definedNames></workbook>",
        )
    return output.getvalue()


def test_constants_and_local_shadowing_survive_reference_only_graph():
    evidence = read_defined_name_evidence(
        package(
            '<definedName name="CompanyName">"Contoso"</definedName>'
            '<definedName name="Limit">5</definedName>'
            '<definedName name="Limit" localSheetId="1">Other!$A$1</definedName>'
        )
    )
    assert evidence["status"] == "complete"
    assert evidence["definitions"][0]["expression"] == '"Contoso"'
    names = relevant_names(evidence, ["companyname", "LIMIT"], "Other")["definitions"]
    assert len(names) == 3
    assert names[-1]["scope"] == "sheet"
    assert names[-1]["scope_sheet"] == "Other"
    assert len(relevant_names(evidence, ["Limit"], "Data")["definitions"]) == 1


def test_qualified_local_name_and_invalid_scope_are_not_global():
    evidence = read_defined_name_evidence(
        package(
            '<definedName name="Rate" localSheetId="0">7</definedName>'
            '<definedName name="Rate" localSheetId="-1">9</definedName>',
            sheets='<sheet name="O\'Brien"/>',
        )
    )
    assert evidence["definitions"][1]["scope"] == "unknown"
    assert (
        len(relevant_names(evidence, ["'O''Brien'!Rate"], "Other")["definitions"]) == 1
    )
    assert relevant_names(evidence, ["Rate"], "Other")["definitions"] == []


def test_workbook_names_can_be_qualified_by_a_known_sheet_case_insensitively():
    evidence = read_defined_name_evidence(
        package(
            '<definedName name="CompanyName">"Contoso"</definedName>'
            '<definedName name="CompanyName" localSheetId="0">"Local"</definedName>'
        )
    )
    names = relevant_names(evidence, ["'data'!COMPANYNAME"], "Other")["definitions"]
    assert {name["scope"] for name in names} == {"sheet", "workbook"}
    assert (
        relevant_names(evidence, ["Missing!CompanyName"], "Data")["definitions"] == []
    )
    assert (
        relevant_names(evidence, ["[1]Data!CompanyName"], "Data")["definitions"] == []
    )


def test_metadata_omissions_are_explicit(monkeypatch):
    import linexcel.doc_evidence as module

    monkeypatch.setattr(module, "MAX_NAME_RECORDS", 1)
    evidence = read_defined_name_evidence(
        package(
            '<definedName name="One">1</definedName>'
            '<definedName name="Two">2</definedName>'
        )
    )
    assert evidence["status"] == "partial"
    assert evidence["omitted"] == 1
    monkeypatch.setattr(module, "MAX_NAME_XML_BYTES", 4)
    assert read_defined_name_evidence(package(""))["status"] == "not_inspected"


@pytest.mark.parametrize("encoding", ["utf-8", "utf-16"])
def test_entity_definitions_are_not_documentation_evidence(encoding):
    data = io.BytesIO()
    with zipfile.ZipFile(data, "w") as z:
        z.writestr(
            "xl/workbook.xml",
            (
                '<!DOCTYPE workbook [<!ENTITY name "injected">]>'
                '<workbook><definedNames><definedName name="Name">'
                "&name;</definedName></definedNames></workbook>"
            ).encode(encoding),
        )
    assert read_defined_name_evidence(data.getvalue())["status"] == "not_inspected"


def test_literal_function_selector_is_sum_and_only_ignores_errors():
    fact = formula_evidence("=AGGREGATE(9,6,A1:A3)")["function_selectors"][0]
    assert fact["operation"] == "SUM"
    assert fact["ignore_errors"]
    assert not fact["ignore_hidden_rows_requested"]
    assert not fact["ignore_nested_aggregates_requested"]


@pytest.mark.parametrize(
    "formula", ["=AGGREGATE(A1,B1,C1:C3)", "=AGGREGATE(TRUE,6,A1:A3)"]
)
def test_nonliteral_or_boolean_selectors_do_not_become_claimed_operations(formula):
    fact = formula_evidence(formula)["function_selectors"][0]
    assert "operation" not in fact


def test_formula_strings_are_not_misclassified_as_references():
    facts = formula_evidence('=IF(A1,"CompanyName",CompanyName)')
    assert facts["references"] == ["A1", "CompanyName"]
    assert formula_evidence('="CompanyName"')["references"] == []


def test_local_bindings_do_not_pick_up_unrelated_workbook_names():
    facts = formula_evidence("=LET(Total,3,Total)+CompanyName")
    assert facts["references"] == ["CompanyName"]
    assert facts["local_scope_references_omitted"]


def test_deep_formula_facts_cannot_abort_the_native_process():
    completed = subprocess.run(
        [
            sys.executable,
            "-c",
            "from linexcel.doc_evidence import formula_evidence; "
            "facts = formula_evidence('=' + '+'.join(['1'] * 3000)); "
            "assert facts['status'] == 'not_inspected'",
        ],
        capture_output=True,
        text=True,
        timeout=15,
    )
    assert completed.returncode == 0, completed.stderr


def test_group_neighbor_value_is_explicitly_representative():
    group = {
        "id": "g",
        "kind": "group",
        "sheet": "Sales",
        "addr": "E4",
        "count": 100,
        "value": 72.5,
        "valueSource": "engine",
    }
    graph = {
        "nodes": [group, {"id": "out", "kind": "cell", "formula": "=SUM(E4:E103)"}],
        "edges": [{"source": "g", "target": "out"}],
    }
    own = build_dossier(graph, "g")
    assert own["representative_cell"] == {"sheet": "Sales", "address": "E4"}
    assert "not an aggregate" in own["value_role"]
    neighbor = build_dossier(graph, "out")["precedents"][0]
    assert neighbor["representative_cell"]["address"] == "E4"
    assert neighbor["group_members"] == 100


def test_missing_graph_name_keeps_source_constant_and_no_invented_self_edge():
    graph = {
        "nodes": [
            {
                "id": "c",
                "kind": "cell",
                "sheet": "Data",
                "formula": "=CompanyName",
                "value": "#NAME?",
            }
        ],
        "edges": [],
        "meta": {
            "definedNameEvidence": read_defined_name_evidence(
                package('<definedName name="CompanyName">"Contoso"</definedName>')
            )
        },
    }
    dossier = build_dossier(graph, "c")
    assert (
        dossier["source_defined_names"]["definitions"][0]["expression"] == '"Contoso"'
    )
    assert not dossier["direct_self_edge_observed"]
    overview = build_workbook_dossier(graph)
    assert overview["source_defined_names"]["definitions"][0]["name"] == "CompanyName"


def test_workbook_name_inventory_is_trimmed_with_counts():
    graph = {
        "nodes": [],
        "edges": [],
        "meta": {
            "definedNameEvidence": {
                "status": "complete",
                "definitions": [
                    {"name": f"Name{i}", "expression": "x" * 1000} for i in range(200)
                ],
            }
        },
    }
    blob = _fit_workbook_dossier(build_workbook_dossier(graph))
    assert len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS
    names = json.loads(blob)["source_defined_names"]
    assert names["omitted_for_size"] + len(names["definitions"]) == 200
    assert len(graph["meta"]["definedNameEvidence"]["definitions"]) == 200
