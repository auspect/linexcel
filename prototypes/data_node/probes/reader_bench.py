"""One reader measurement in a fresh process; prints one JSON line.

usage: reader_bench.py <reader> <mode> <xlsx> <sheet> <header_row 0-based> [col,col]
readers: calamine | fastexcel | polars | formualizer | stream_parse
modes:   full | head1000 | cols2 | table:<name> | query
"""

import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from winmem import peak_mb  # noqa: E402

reader, mode, path, sheet = sys.argv[1:5]
header_row = int(sys.argv[5]) if len(sys.argv) > 5 else 0
cols = sys.argv[6].split(",") if len(sys.argv) > 6 else []
if cols and cols[0].startswith("idx:"):
    cols = [int(c.removeprefix("idx:")) for c in cols]
out = {"reader": reader, "mode": mode, "file": Path(path).name, "sheet": sheet}
t0 = time.perf_counter()


def done(**kw):
    out.update(kw)
    out["seconds"] = round(time.perf_counter() - t0, 2)
    out.update(peak_mb())
    print(json.dumps(out, ensure_ascii=False, default=str), flush=True)


try:
    if reader == "calamine":
        import python_calamine as pc

        wb = pc.CalamineWorkbook.from_path(path, load_tables=mode.startswith("table:"))
        if mode.startswith("table:"):
            table = wb.get_table_by_name(mode.split(":", 1)[1])
            rows = table.to_python()
            out["tableColumns"] = getattr(table, "columns", None)
        else:
            s = wb.get_sheet_by_name(sheet)
            out["openSeconds"] = round(time.perf_counter() - t0, 2)
            rows = s.to_python(nrows=1000 + header_row if mode == "head1000" else None)
            if mode == "cols2":
                header = rows[header_row]
                idx = [c if isinstance(c, int) else header.index(c) for c in cols]
                rows = [[r[i] for i in idx] for r in rows]
        types = {}
        for r in rows[header_row + 1 : header_row + 1001]:
            for i, v in enumerate(r):
                types.setdefault(i, set()).add(type(v).__name__)
        done(
            rows=len(rows),
            width=max((len(r) for r in rows), default=0),
            header=[str(v) for v in rows[header_row]][:12] if rows else None,
            pythonTypesFirst1000={k: sorted(v) for k, v in list(types.items())[:12]},
        )
    elif reader in {"fastexcel", "polars"}:
        import fastexcel
        import polars as pl

        if reader == "fastexcel":
            xl = fastexcel.read_excel(path)
            kw = {"header_row": header_row}
            if mode == "head1000":
                kw["n_rows"] = 1000
            if mode == "cols2":
                kw["use_columns"] = cols
            if mode.startswith("table:"):
                obj = xl.load_table(mode.split(":", 1)[1])
            else:
                obj = xl.load_sheet(sheet, **kw)
            out["loadSeconds"] = round(time.perf_counter() - t0, 2)
            frame = obj.to_polars()
            done(
                rows=frame.height,
                width=frame.width,
                schema={c.name: c.dtype for c in obj.selected_columns[:60]},
                dtypeFrom=sorted({c.dtype_from for c in obj.selected_columns}),
                nullCounts={
                    k: v
                    for k, v in list(frame.null_count().row(0, named=True).items())[:60]
                },
            )
        else:
            read_options = {"header_row": header_row}
            if mode == "head1000":
                read_options["n_rows"] = 1000
            if mode == "cols2":
                read_options["use_columns"] = cols
            frame = pl.read_excel(
                path, sheet_name=sheet, engine="calamine", read_options=read_options
            )
            out["readSeconds"] = round(time.perf_counter() - t0, 2)
            if mode == "query" and cols:
                key, value = cols
                t = time.perf_counter()
                result = (
                    frame.lazy()
                    .filter(pl.col(value).is_not_null())
                    .group_by(key)
                    .agg(pl.col(value).sum().alias("total"), pl.len().alias("n"))
                    .sort("total", descending=True)
                    .limit(5)
                    .collect()
                )
                out["querySeconds"] = round(time.perf_counter() - t, 4)
                out["queryTop"] = result.rows()[:3]
            done(
                rows=frame.height,
                width=frame.width,
                schema={k: str(v) for k, v in list(frame.schema.items())[:60]},
            )
    elif reader == "formualizer":
        import formualizer as fz

        wb = fz.load_workbook(path)
        done(sheets=wb.sheet_names if hasattr(wb, "sheet_names") else None)
    elif reader == "stream_parse":
        # linexcel's reading primitive (ZIP + ElementTree iterparse, detach rows)
        # with no storage at all: the floor of today's streaming import.
        import zipfile
        from xml.etree import ElementTree as ET

        from linexcel.formula_source import _worksheet_parts

        with zipfile.ZipFile(path) as archive:
            part = next(p for n, p in _worksheet_parts(archive) if n == sheet)
            cells = 0
            with archive.open(part) as source:
                stack = []
                for event, element in ET.iterparse(source, events=("start", "end")):
                    if event == "start":
                        stack.append(element)
                        continue
                    tag = element.tag.rsplit("}", 1)[-1]
                    if tag == "c":
                        element.findtext("{*}v")
                        cells += 1
                        stack[-2].remove(element)
                        element.clear()
                    elif tag == "row":
                        stack[-2].remove(element)
                        element.clear()
                    stack.pop()
        done(cells=cells)
    else:
        raise SystemExit(f"unknown reader {reader}")
except BaseException as exc:
    done(status="failed", error=f"{type(exc).__name__}: {exc}"[:500])
