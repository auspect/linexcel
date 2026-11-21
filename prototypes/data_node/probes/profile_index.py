"""cProfile of serve's build_index on one workbook; writes the top functions.

usage: profile_index.py <xlsx> <index.sqlite> <out.json>
"""

import cProfile
import io
import json
import pstats
import sys
import time
from pathlib import Path

from linexcel.lazy_store import build_index

xlsx, index, out = map(Path, sys.argv[1:4])
index.unlink(missing_ok=True)
profile = cProfile.Profile()
started = time.perf_counter()
profile.runcall(build_index, xlsx, index)
seconds = time.perf_counter() - started
stats = pstats.Stats(profile)
rows = []
for (filename, line, name), (cc, nc, tt, ct, _callers) in stats.stats.items():
    rows.append(
        {
            "function": f"{Path(filename).name}:{line}({name})",
            "calls": nc,
            "tottime": round(tt, 2),
            "cumtime": round(ct, 2),
        }
    )
top_cum = sorted(rows, key=lambda r: -r["cumtime"])[:30]
top_tot = sorted(rows, key=lambda r: -r["tottime"])[:20]
result = {
    "workbook": xlsx.name,
    "secondsUnderProfiler": round(seconds, 1),
    "topCumulative": top_cum,
    "topTotal": top_tot,
}
out.write_text(json.dumps(result, indent=1), "utf-8")
buffer = io.StringIO()
pstats.Stats(profile, stream=buffer).sort_stats("cumulative").print_stats(30)
print(buffer.getvalue()[-6000:])
