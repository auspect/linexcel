"""Workbook context extraction and optional Linux screenshot rendering."""

from __future__ import annotations

import datetime
import io
import math
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from linexcel.refs import num_to_col

PREVIEW_ROWS = 12
PREVIEW_COLUMNS = 8
MAX_COMMENTS_PER_SHEET = 20
MAX_COMMENT_SCAN_CELLS = 100_000
MAX_COMMENT_CHARS = 1_000
MAX_MERGED_RANGES = 30


class WorkbookRenderError(RuntimeError):
    """Raised when the optional workbook screenshot renderer is unavailable."""


def extract_workbook_context(
    data: bytes,
    filename: str = "workbook.xlsx",
    *,
    preview_rows: int = PREVIEW_ROWS,
    preview_columns: int = PREVIEW_COLUMNS,
) -> dict[str, Any]:
    """Extract bounded presentation context without launching Excel.

    The preview preserves the first cells as they are, without guessing which
    row is a header. Comments and sheet layout markers complement formula
    lineage with the cues users usually see when opening a workbook.
    """
    workbook = load_workbook(
        io.BytesIO(data),
        read_only=False,
        data_only=False,
        keep_vba=filename.lower().endswith((".xlsm", ".xltm")),
    )
    warnings: list[str] = []
    sheets: list[dict[str, Any]] = []
    total_comments = 0
    try:
        for worksheet in workbook.worksheets:
            max_row = max(worksheet.max_row or 1, 1)
            max_column = max(worksheet.max_column or 1, 1)
            row_limit = min(max_row, preview_rows)
            column_limit = min(max_column, preview_columns)
            preview = [
                {
                    "row": row[0].row,
                    "values": [_safe_value(cell.value) for cell in row],
                }
                for row in worksheet.iter_rows(
                    min_row=1,
                    max_row=row_limit,
                    min_col=1,
                    max_col=column_limit,
                )
            ]
            comments, comments_truncated = _extract_comments(
                worksheet, max_row, max_column
            )
            total_comments += len(comments)
            if comments_truncated:
                warnings.append(
                    f"Comments on '{worksheet.title}' were truncated for inspection"
                )
            sheets.append(
                {
                    "name": worksheet.title,
                    "visibility": worksheet.sheet_state,
                    "dimensions": {"rows": max_row, "columns": max_column},
                    "preview_range": f"A1:{num_to_col(column_limit)}{row_limit}",
                    "preview": preview,
                    "freeze_panes": str(worksheet.freeze_panes)
                    if worksheet.freeze_panes
                    else None,
                    "merged_ranges": [
                        str(cell_range)
                        for cell_range in list(worksheet.merged_cells.ranges)[
                            :MAX_MERGED_RANGES
                        ]
                    ],
                    "hidden_columns": _hidden_columns(worksheet, column_limit),
                    "comments": comments,
                }
            )
    finally:
        workbook.close()
    return {
        "filename": filename,
        "sheets": sheets,
        "stats": {"sheets": len(sheets), "comments": total_comments},
        "warnings": warnings,
    }


def render_workbook_screenshots(
    data: bytes,
    filename: str,
    output_dir: str | Path,
    *,
    dpi: int = 144,
    timeout: int = 60,
) -> list[Path]:
    """Render workbook pages to PNG with LibreOffice and Poppler on Linux.

    LibreOffice runs headlessly; no desktop Excel process is needed. It exports
    the workbook to PDF, then ``pdftoppm`` creates one PNG per rendered page.
    """
    office = shutil.which("libreoffice") or shutil.which("soffice")
    converter = shutil.which("pdftoppm")
    if not office or not converter:
        raise WorkbookRenderError(
            "Workbook screenshots require LibreOffice and pdftoppm. "
            "Install LibreOffice and poppler-utils, then try again."
        )
    if dpi <= 0:
        raise ValueError("dpi must be positive")
    if timeout <= 0:
        raise ValueError("timeout must be positive")

    target = Path(output_dir)
    target.mkdir(parents=True, exist_ok=True)
    suffix = Path(filename).suffix.lower()
    if suffix not in {".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm"}:
        suffix = ".xlsx"
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "-", Path(filename).stem).strip(".-")
    stem = stem or "workbook"

    with tempfile.TemporaryDirectory(prefix="linexcel-render-") as temp_dir:
        temp = Path(temp_dir)
        input_path = temp / f"workbook{suffix}"
        pdf_dir = temp / "pdf"
        input_path.write_bytes(data)
        pdf_dir.mkdir()
        try:
            subprocess.run(
                [
                    office,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(pdf_dir),
                    str(input_path),
                ],
                check=True,
                capture_output=True,
                text=True,
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as exc:
            raise WorkbookRenderError(
                f"LibreOffice did not finish within {timeout} seconds"
            ) from exc
        except subprocess.CalledProcessError as exc:
            details = (exc.stderr or exc.stdout or "unknown error").strip()
            raise WorkbookRenderError(
                f"LibreOffice could not render the workbook: {details}"
            ) from exc

        pdfs = list(pdf_dir.glob("*.pdf"))
        if not pdfs:
            raise WorkbookRenderError("LibreOffice did not produce a PDF")
        prefix = target / stem
        try:
            subprocess.run(
                [converter, "-png", "-r", str(dpi), str(pdfs[0]), str(prefix)],
                check=True,
                capture_output=True,
                text=True,
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as exc:
            raise WorkbookRenderError(
                f"PDF conversion did not finish within {timeout} seconds"
            ) from exc
        except subprocess.CalledProcessError as exc:
            details = (exc.stderr or exc.stdout or "unknown error").strip()
            raise WorkbookRenderError(
                f"pdftoppm could not create screenshots: {details}"
            ) from exc
    screenshots = sorted(target.glob(f"{stem}-*.png"))
    if not screenshots:
        raise WorkbookRenderError("pdftoppm did not produce PNG screenshots")
    return screenshots


def _safe_value(value: Any) -> Any:
    if isinstance(value, (datetime.date, datetime.datetime, datetime.time)):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _extract_comments(
    worksheet, max_row: int, max_column: int
) -> tuple[list[dict], bool]:
    scan_rows = min(max_row, math.ceil(MAX_COMMENT_SCAN_CELLS / max_column))
    comments: list[dict[str, str | None]] = []
    for row in worksheet.iter_rows(
        min_row=1, max_row=scan_rows, min_col=1, max_col=max_column
    ):
        for cell in row:
            if cell.comment is None:
                continue
            comments.append(
                {
                    "cell": cell.coordinate,
                    "author": cell.comment.author,
                    "text": cell.comment.text[:MAX_COMMENT_CHARS],
                }
            )
            if len(comments) >= MAX_COMMENTS_PER_SHEET:
                return comments, True
    return comments, max_row > scan_rows


def _hidden_columns(worksheet, column_limit: int) -> list[str]:
    return [
        num_to_col(column)
        for column in range(1, column_limit + 1)
        if worksheet.column_dimensions[num_to_col(column)].hidden
    ]
