"""Current serve views on a built index, timed in a fresh process.

usage:
  probe_views.py presentation <xlsx> <index> <sheet> <row> <out.json>
  probe_views.py build_graph <index> <hierarchy.sqlite> <out.json>
  probe_views.py datasette <index> <source-data.sqlite> <out.json>
"""

import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from winmem import peak_mb  # noqa: E402

command, *args = sys.argv[1:]
out = Path(args[-1])
result = {"command": command, "args": [Path(a).name for a in args[:-1]]}
started = time.perf_counter()
try:
    if command == "presentation":
        from linexcel.excel_presentation import sheet_presentation

        xlsx, index, sheet, row = args[:4]
        view = sheet_presentation(
            Path(xlsx), Path(index), sheet, row=int(row), rows=60, columns=20
        )
        result.update(status="ok", cells=len(view["cells"]), extent=view["extent"])
    elif command == "build_graph":
        from linexcel.graph_store import build_graph

        index, output = map(Path, args[:2])
        output.unlink(missing_ok=True)
        meta = build_graph(index, output)
        result.update(
            status="ok",
            sizeMB=round(output.stat().st_size / 1e6, 1),
            summary={k: v for k, v in meta.items() if isinstance(v, (int, float, str))},
        )
    elif command == "datasette":
        from linexcel.web_data import build_source_database

        index, output = map(Path, args[:2])
        output.unlink(missing_ok=True)
        meta = build_source_database(index, output, max_bytes=10**13)
        result.update(
            status="ok",
            sizeMB=round(output.stat().st_size / 1e6, 1),
            sheets=[(s["sheet"], s["layout"], s["cellCount"]) for s in meta["sheets"]][
                :12
            ],
        )
except BaseException as exc:
    result.update(status="failed", error=f"{type(exc).__name__}: {exc}"[:400])
result["seconds"] = round(time.perf_counter() - started, 2)
result.update(peak_mb())
out.write_text(json.dumps(result, indent=1, default=str), "utf-8")
print(json.dumps(result, default=str))
