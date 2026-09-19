"""High-level documentation retains qualifications and inspectable raw evidence."""

from copy import deepcopy

import pytest

from linexcel.result import LineageResult


@pytest.mark.parametrize("overview", [False, True])
def test_result_exports_qualified_prose_and_preserves_source(overview):
    graph = {
        "nodes": [{"id": "c:S!B1", "kind": "cell", "formula": "=SUM(A1:A2)"}],
        "edges": [],
        "meta": {},
    }
    original = deepcopy(graph)
    result = LineageResult(graph, None)
    raw = "Source formula: `=SUM(A1:A3)`"
    evidence = {}
    options = {"provider": lambda *args, **kwargs: raw, "validation_results": evidence}
    if overview:
        text = result.document_workbook(**options)
        html = result.to_html(workbook_doc=text)
        report = evidence["workbook"]
    else:
        docs = result.document(**options)
        text = docs["c:S!B1"]
        html = result.to_html(docs=docs)
        report = evidence["c:S!B1"]
    assert text.startswith("> ")
    assert text.endswith(raw)
    assert report["status"] == "qualified"
    assert report["raw_response"] == raw
    assert "SUM(A1:A3)" in html
    assert graph == original
