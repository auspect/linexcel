"""Excel workbook data lineage analysis.

Pipeline: formula extraction (formualizer, Rust engine), then stretched pattern
grouping (R1C1 canonicalization), then dependency graph (cells, ranges, defined
names, VBA), then composite function decomposition with step-by-step evaluation,
and finally optional AI documentation, from the provider of your choice,
grounded in deterministic lineage.

Usable as a standalone library (marimo, Jupyter, scripts), without
a FastAPI server or AI key:

    from linexcel import analyze
    result = analyze("workbook.xlsx")   # -> LineageResult
    result                               # interactive graph in marimo
    result.save_html("lineage.html")     # standalone offline viewer
    result.stats, result.warnings        # metadata
    result.document(base_url=...)        # AI documentation (optional, opt-in)
"""

from importlib.metadata import PackageNotFoundError as _PackageNotFoundError
from importlib.metadata import version as _installed_version

from linexcel.analyzer import analyze_workbook
from linexcel.execution import ExecutionPolicy
from linexcel.insights import WorkbookRenderError
from linexcel.result import LineageResult, analyze, read_macros

try:
    __version__ = _installed_version("linexcel")
except _PackageNotFoundError:  # pragma: no cover - source tree without install
    __version__ = "0.0.0.dev0"

__all__ = [
    "analyze",
    "read_macros",
    "ExecutionPolicy",
    "LineageResult",
    "WorkbookRenderError",
    "analyze_workbook",
]
