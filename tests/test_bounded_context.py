"""Optional report context must not materialize millions of formatted cells."""

import io
import re
import zipfile

from openpyxl import Workbook
from openpyxl.comments import Comment

from linexcel import insights
from linexcel.aidoc import build_workbook_dossier


def test_large_context_streams_previews_and_reports_omissions(monkeypatch):
    workbook = Workbook()
    sheet = workbook.active
    sheet["B2"] = "Revenue"
    sheet["C2"] = "=10+20"
    sheet["B2"].comment = Comment("Private context", "Analyst")
    sheet.freeze_panes = "B2"
    output = io.BytesIO()
    workbook.save(output)
    original = insights.load_workbook
    modes = []

    def load(*args, **kwargs):
        modes.append(kwargs["read_only"])
        return original(*args, **kwargs)

    monkeypatch.setattr(insights, "load_workbook", load)
    monkeypatch.setattr(insights, "MAX_CONTEXT_XML_BYTES", 0)
    context = insights.extract_workbook_context(output.getvalue())
    assert modes == [True]
    assert context["sheets"][0]["preview"][1] == {
        "row": 2,
        "values": [None, "Revenue", "=10+20"],
    }
    assert context["sheets"][0]["comments"] == []
    assert "not scanned" in context["warnings"][0]
    dossier = build_workbook_dossier({"nodes": [], "meta": {}}, context=context)
    assert dossier["warnings"] == context["warnings"]


def test_missing_dimensions_still_reads_bounded_preview(monkeypatch):
    workbook = Workbook()
    workbook.active["C4"] = "Visible without dimensions"
    output = io.BytesIO()
    workbook.save(output)
    patched = io.BytesIO()
    with zipfile.ZipFile(output) as source, zipfile.ZipFile(patched, "w") as target:
        for part in source.infolist():
            data = source.read(part)
            if part.filename == "xl/worksheets/sheet1.xml":
                data = re.sub(rb"<dimension\b[^>]*/>", b"", data)
            target.writestr(part, data)
    monkeypatch.setattr(insights, "MAX_CONTEXT_XML_BYTES", 0)
    context = insights.extract_workbook_context(patched.getvalue())
    sheet = context["sheets"][0]
    assert sheet["preview"][3]["values"][2] == "Visible without dimensions"
    assert sheet["dimensions"] == {"rows": None, "columns": None}
