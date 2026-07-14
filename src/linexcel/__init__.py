"""Excel workbook data lineage analysis.

Pipeline: formula extraction (formualizer, Rust engine) →
stretched pattern grouping (R1C1 canonicalization) →
dependency graph (cells, ranges, defined names, VBA) →
composite function decomposition with step-by-step evaluation →
optional AI documentation (google-genai) grounded in deterministic lineage.

Usable as a standalone library (marimo, Jupyter, scripts), without
a FastAPI server or AI key:

    from linexcel import analyze
    result = analyze("workbook.xlsx")   # -> LineageResult
    result                               # interactive graph in marimo
    result.save_html("lineage.html")     # standalone offline viewer
    result.stats, result.warnings        # metadata
    result.document(api_key="...")       # AI documentation (optional)
"""

from linexcel.analyzer import analyze_workbook
from linexcel.insights import WorkbookRenderError
from linexcel.result import LineageResult, analyze

__all__ = [
    "analyze",
    "LineageResult",
    "WorkbookRenderError",
    "analyze_workbook",
]
