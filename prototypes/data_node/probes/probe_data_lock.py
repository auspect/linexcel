"""Does the first Data-tab request block the whole server?

Imports a workbook, then requests /data (Datasette snapshot build) in one
thread while another thread polls light endpoints every 0.25 s and records
their latency.
usage: probe_data_lock.py <xlsx> <data-dir> <out.json>
"""

import hashlib
import io
import json
import sys
import threading
import time
from pathlib import Path

from linexcel.web import create_app

workbook, data_dir, out = sys.argv[1:4]
app = create_app(Path(data_dir), resolve_user=lambda env: env.get("user"))
owner = hashlib.sha256(b"alice").hexdigest()


def call(method, path, query=""):
    status = []
    started = time.perf_counter()
    raw = b"".join(
        app(
            {
                "REQUEST_METHOD": method,
                "PATH_INFO": path,
                "QUERY_STRING": query,
                "CONTENT_TYPE": "application/json",
                "CONTENT_LENGTH": "0",
                "wsgi.input": io.BytesIO(b""),
                "user": "alice",
            },
            lambda s, h: status.append(s),
        )
    )
    return int(status[0].split()[0]), time.perf_counter() - started, raw


created = app.store.uploads.local(
    owner,
    {
        "path": str(Path(workbook).resolve()),
        "budget": {"seconds": 900, "memoryMb": 1024},
    },
)
project = created["project"]["id"]
while app.store.get_task(owner, created["task"]["id"])["status"] not in {
    "succeeded",
    "failed",
}:
    time.sleep(0.5)
result = {
    "workbook": Path(workbook).name,
    "import": app.store.get_task(owner, created["task"]["id"])["status"],
}
latencies, done = [], threading.Event()


def poll():
    while not done.is_set():
        code, seconds, _ = call("GET", "/api/projects")
        latencies.append((round(time.perf_counter() - t0, 2), code, round(seconds, 3)))
        time.sleep(0.25)


t0 = time.perf_counter()
poller = threading.Thread(target=poll)
poller.start()
time.sleep(1)
code, seconds, raw = call("GET", f"/api/projects/{project}/data")
done.set()
poller.join()
result["dataTab"] = {"code": code, "seconds": round(seconds, 2)}
result["pollCount"] = len(latencies)
result["maxPollLatency"] = max(poll[2] for poll in latencies)
result["pollsSlowerThan1s"] = [poll for poll in latencies if poll[2] > 1]
app.close()
Path(out).write_text(json.dumps(result, indent=1), "utf-8")
print(json.dumps(result))
