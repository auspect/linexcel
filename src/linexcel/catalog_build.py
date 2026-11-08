"""Fair, bounded generation construction, separate from catalog publication."""

from __future__ import annotations

import hashlib
import json
from collections import deque
from typing import Any

from linexcel.header_bands import HeaderBands, numeric
from linexcel.lazy_store import _unpack
from linexcel.refs import a1


def fair_budgets(capacities, total):
    """Water fill: small sheets finish, remaining sheets share unused capacity."""
    result = {key: 0 for key in capacities}
    active = sorted(key for key, count in capacities.items() if count)
    while total and active:
        share, extra = divmod(total, len(active))
        spent = 0
        for i, key in enumerate(active):
            amount = min(capacities[key] - result[key], share + (i < extra))
            result[key] += amount
            spent += amount
        total -= spent
        active = [k for k in active if result[k] < capacities[k]]
    return result


def schema(db, fts5):
    db.execute(
        "CREATE TABLE entries(id INTEGER PRIMARY KEY,project TEXT,kind TEXT,"
        "node TEXT,sheet TEXT,address TEXT,text TEXT,labels TEXT,name TEXT,"
        "occurrence_count INTEGER,occurrences TEXT,count_complete INTEGER)"
    )
    if fts5:
        db.execute(
            "CREATE VIRTUAL TABLE search USING fts5(text,content='entries',"
            "content_rowid='id',tokenize='unicode61 remove_diacritics 2')"
        )
    db.execute("CREATE INDEX scope ON entries(project,kind,id)")
    db.execute(
        "CREATE TABLE numeric_cells(id INTEGER PRIMARY KEY,entry INTEGER,"
        "node TEXT,address TEXT,number REAL,labels TEXT)"
    )
    db.execute("CREATE INDEX numeric ON numeric_cells(number,entry)")
    db.execute(
        "CREATE TABLE header_bands(project TEXT,sheet TEXT,data TEXT,coverage TEXT)"
    )


def coverage(total):
    return dict(
        nodesTotal=total,
        nodesScanned=0,
        nodesUnscanned=total,
        entries=0,
        eligibleObserved=0,
        numericObserved=0,
        numericAssessed=0,
        numericLabelled=0,
        numericIndexed=0,
        numericUnassessed=0,
        textTruncated=0,
        capsHit=[],
    )


class WorkbookBuilder:
    def __init__(self, db, index, item, limits, fts5, expired, progress=lambda _: None):
        # Import at call time to keep publication and construction independent.
        from linexcel.catalog import _json, normalize

        self.json, self.normalize = _json, normalize
        self.db, self.index, self.item = db, index, item
        self.limits, self.fts5, self.expired = limits, fts5, expired
        self.progress = progress
        self.disk_full = False
        self.counts = item["coverage"] = coverage(item["nodes"])
        inventory = index.db.execute(
            "SELECT sheet,kind,count(*) FROM nodes GROUP BY sheet,kind"
        ).fetchall()
        names = [
            s if isinstance(s, str) else s["name"] for s in index.meta.get("sheets", [])
        ]
        self.sheet_names = set(names)
        self.sheets: dict[str | None, dict[str, Any]] = {
            name: {**coverage(0), "sheet": name, "byKind": {}, "nodeKinds": {}}
            for name in names
        }
        for sheet, kind, count in inventory:
            scope = self.sheets.setdefault(
                sheet, {**coverage(0), "sheet": sheet, "byKind": {}, "nodeKinds": {}}
            )
            scope["nodesTotal"] += count
            scope["nodeKinds"][kind] = count
        self.tables = {}
        for (blob,) in index.db.execute(
            "SELECT data FROM table_definitions ORDER BY name_key,sheet"
        ):
            table = json.loads(blob)
            self.tables.setdefault(table["sheet"], []).append(table)
        self.maps = {}
        self.groups = {sheet: {} for sheet in self.sheets}

    def cap(self, scope, cap):
        if cap not in scope["capsHit"]:
            scope["capsHit"].append(cap)

    def transaction(self, scope, operation):
        if self.disk_full:
            self.cap(scope, "catalog_bytes")
            return None
        try:
            self.db.execute("SAVEPOINT entry")
            result = operation()
            self.db.execute("RELEASE entry")
            return result
        except Exception as error:
            if "full" not in str(error):
                raise
            if self.db.in_transaction:
                self.db.rollback()
            self.cap(scope, "catalog_bytes")
            self.disk_full = True
            return None

    def add(self, kind, sheet, address, text, *, node=None, labels=(), body=False):
        scope = self.sheets[sheet]
        counts = scope["byKind"].setdefault(kind, {"eligibleObserved": 0, "entries": 0})
        counts["eligibleObserved"] += 1
        scope["eligibleObserved"] += 1
        # Keep original, complete normalized text identity even if its searchable
        # representation is truncated. Different long strings must not coalesce.
        key = hashlib.sha256(self.normalize(text).encode()).digest()
        groups = self.groups[sheet]
        if body and key in groups:
            record = groups[key]
            record["count"] += 1
            if len(record["occurrences"]) < self.limits.text_occurrences:
                record["occurrences"].append({"node": node, "address": address})
            return record["id"]
        if self.counts["entries"] >= self.limits.entries_per_workbook:
            self.cap(scope, "entries_per_workbook")
            return None
        if body and scope["bodyEntries"] >= scope["entryBudget"]:
            self.cap(scope, "entries_per_sheet")
            return None
        if self.expired():
            self.cap(scope, "build_seconds")
            return None
        if len(text) > self.limits.text_chars:
            scope["textTruncated"] += 1
            self.cap(scope, "text_chars")
        occurrences = [{"node": node, "address": address}] if body else []

        def write():
            cursor = self.db.execute(
                "INSERT INTO entries(project,kind,node,sheet,address,text,labels,"
                "name,occurrence_count,occurrences,count_complete) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (
                    self.item["projectId"],
                    kind,
                    node,
                    sheet,
                    address,
                    self.normalize(text[: self.limits.text_chars]),
                    self.json(labels),
                    text if kind != "cell" else None,
                    1,
                    self.json(occurrences),
                    0 if body else 1,
                ),
            )
            if self.fts5:
                self.db.execute(
                    "INSERT INTO search(rowid,text) VALUES (?,?)",
                    (cursor.lastrowid, self.normalize(text[: self.limits.text_chars])),
                )
            return cursor.lastrowid

        entry = self.transaction(scope, write)
        if entry is None:
            return None
        self.counts["entries"] += 1
        scope["entries"] += 1
        counts["entries"] += 1
        if body:
            scope["bodyEntries"] += 1
            groups[key] = {
                "id": entry,
                "count": 1,
                "occurrences": occurrences,
                "labels": False,
            }
        return entry

    def structural(self):
        # All sheets before tables, names, or inferred headers, even at tiny caps.
        for sheet in self.sheets:
            if sheet in self.sheet_names:
                self.add("sheet", sheet, None, sheet)
        for sheet, tables in self.tables.items():
            for table in tables:
                r1, c1, r2, c2 = table["bounds"]
                self.add("table", sheet, f"{a1(r1, c1)}:{a1(r2, c2)}", table["name"])
                for i, header in enumerate(table.get("headers", [])):
                    self.add(
                        "column", sheet, a1(r1, c1 + i), f"{table['name']} {header}"
                    )
        for node_id, sheet, blob in self.index.db.execute(
            "SELECT id,sheet,data FROM nodes WHERE kind='name' ORDER BY seq"
        ):
            node = _unpack(blob)
            self.sheets[sheet]["nodesScanned"] += 1
            self.add(
                "name", sheet, node.get("cell"), node.get("name", node_id), node=node_id
            )
        for sheet, scope in self.sheets.items():
            if sheet not in self.sheet_names:
                continue
            bands = HeaderBands(
                self.index, sheet, self.tables.get(sheet, []), self.limits, self.expired
            )
            self.maps[sheet] = bands
            scope["headerBands"] = bands.coverage
            for cap in bands.coverage["capsHit"]:
                self.cap(scope, cap)
            self.transaction(
                scope,
                lambda: self.db.execute(
                    "INSERT INTO header_bands VALUES (?,?,?,?)",
                    (
                        self.item["projectId"],
                        sheet,
                        self.json(bands.records()),
                        self.json(bands.coverage),
                    ),
                ),
            )
            for band in bands.bands:
                for column, text in band.headers.items():
                    self.add("column", sheet, a1(band.header_row, column), text)

    def labels(self, sheet, row, column):
        result = self.index.cell_labels(sheet, row, column)
        inferred = self.maps[sheet].labels(row, column) if sheet in self.maps else []
        sources = {(label["sheet"], label["cell"], label["role"]) for label in result}
        for label in inferred:
            if label["role"] == "column" and any(
                existing["method"] == "table_column" for existing in result
            ):
                continue
            key = (label["sheet"], label["cell"], label["role"])
            if key not in sources:
                result.append(label)
                sources.add(key)
        return result

    def body(self):
        capacity = {
            sheet: scope["nodesTotal"] - scope["nodeKinds"].get("name", 0)
            for sheet, scope in self.sheets.items()
        }
        # Null (workbook-scope) is sorted explicitly, never compared with a str.
        ordered = {str(i): value for i, value in enumerate(capacity.values())}
        entries = fair_budgets(
            ordered, max(0, self.limits.entries_per_workbook - self.counts["entries"])
        )
        numerics = fair_budgets(ordered, self.limits.numeric_cells_per_workbook)
        queue = deque()
        for i, (sheet, scope) in enumerate(self.sheets.items()):
            scope.update(
                entryBudget=entries[str(i)],
                numericBudget=numerics[str(i)],
                bodyEntries=0,
                textCellsObserved=0,
                textCellsOmitted=0,
            )
            cursor = self.index.db.execute(
                "SELECT id,kind,row,col,data FROM nodes WHERE sheet IS ? "
                "AND kind!='name' ORDER BY row,col,seq",
                (sheet,),
            )
            queue.append((sheet, scope, cursor, 0))
        while queue and not self.expired() and not self.disk_full:
            sheet, scope, cursor, ordinal = queue.popleft()
            rows = cursor.fetchmany(512)
            for node_id, kind, row, column, blob in rows:
                scope["nodesScanned"] += 1
                ordinal += 1
                if (
                    kind not in {"cell", "input", "array_result"}
                    or not row
                    or not column
                ):
                    continue
                node = _unpack(blob)
                value = node.get("value")
                if isinstance(value, str) and value.strip():
                    scope["textCellsObserved"] += 1
                    entry = self.add(
                        "cell", sheet, node.get("cell"), value, node=node_id, body=True
                    )
                    if entry is None:
                        scope["textCellsOmitted"] += 1
                    continue
                is_numeric = numeric(node)
                if is_numeric:
                    scope["numericObserved"] += 1
                # Systematic, sheet-local sampling spans the entire sheet; this
                # does not stop at the first N numeric cells. The bound is on
                # assessments, with the exact unassessed population reported.
                size = max(1, capacity[sheet])
                budget = scope["numericBudget"]
                selected = (ordinal * budget // size) != (
                    (ordinal - 1) * budget // size
                )
                if not selected:
                    continue
                if is_numeric:
                    scope["numericAssessed"] += 1
                labels = self.labels(sheet, row, column)
                if is_numeric and labels:
                    scope["numericLabelled"] += 1
                if not labels:
                    continue
                text = " ".join(label["text"] for label in labels)
                entry = self.add(
                    "cell",
                    sheet,
                    node.get("cell"),
                    text,
                    node=node_id,
                    labels=labels,
                    body=True,
                )
                if entry is not None:
                    key = hashlib.sha256(self.normalize(text).encode()).digest()
                    self.groups[sheet][key]["labels"] = True
                    if is_numeric:
                        saved = self.transaction(
                            scope,
                            lambda: self.db.execute(
                                "INSERT INTO numeric_cells(entry,node,address,"
                                "number,labels) "
                                "VALUES (?,?,?,?,?)",
                                (
                                    entry,
                                    node_id,
                                    node.get("cell"),
                                    float(value),
                                    self.json(labels),
                                ),
                            ),
                        )
                        if saved is not None:
                            scope["numericIndexed"] += 1
            if rows:
                queue.append((sheet, scope, cursor, ordinal))
            elif queue:
                # Sheets with few distinct strings return their unused share.
                # Every unfinished sheet receives the same increment; a large
                # early sheet cannot take capacity reserved for later sheets.
                unused = scope["entryBudget"] - scope["bodyEntries"]
                share, extra = divmod(unused, len(queue))
                for i, (_, other, _, _) in enumerate(queue):
                    other["entryBudget"] += share + (i < extra)
            self.progress(sum(s["nodesScanned"] for s in self.sheets.values()))
        for sheet, scope in self.sheets.items():
            scope["sourceSheet"] = sheet in self.sheet_names
            scope["nodesUnscanned"] = scope["nodesTotal"] - scope["nodesScanned"]
            scope["numericUnassessed"] = (
                scope["numericObserved"] - scope["numericAssessed"]
            )
            if scope["numericUnassessed"]:
                self.cap(scope, "numeric_cells_per_workbook")
            if scope["nodesUnscanned"]:
                self.cap(scope, "catalog_bytes" if self.disk_full else "build_seconds")
            scope["byKind"]["text"] = {
                "cellsObserved": scope["textCellsObserved"],
                "cellsOmitted": scope["textCellsOmitted"],
                "populationComplete": not scope["nodesUnscanned"],
            }
            scope["byKind"]["numeric"] = {
                "cellsObserved": scope["numericObserved"],
                "cellsAssessed": scope["numericAssessed"],
                "cellsLabelled": scope["numericLabelled"],
                "cellsIndexed": scope["numericIndexed"],
                "cellsUnassessed": scope["numericUnassessed"],
                "populationComplete": not scope["nodesUnscanned"],
            }

            def finalize():
                # One transaction per sheet, rather than a filesystem commit
                # for every distinct string. Initial flags are false, so a
                # disk-full rollback cannot claim an unflushed count is exact.
                self.db.execute(
                    "UPDATE entries SET count_complete=? WHERE project=? "
                    "AND sheet IS ? AND kind='cell'",
                    (int(not scope["nodesUnscanned"]), self.item["projectId"], sheet),
                )
                self.db.executemany(
                    "UPDATE entries SET occurrence_count=?,occurrences=?,"
                    "count_complete=? WHERE id=?",
                    (
                        (
                            record["count"],
                            self.json(record["occurrences"]),
                            int(
                                not scope["nodesUnscanned"]
                                and not (
                                    record["labels"] and scope["numericUnassessed"]
                                )
                            ),
                            record["id"],
                        )
                        for record in self.groups[sheet].values()
                        if record["count"] > 1
                        or record["labels"]
                        and scope["numericUnassessed"]
                    ),
                )
                return True

            self.transaction(scope, finalize)
            scope["capsHit"] = sorted(scope["capsHit"])
        self.counts["sheets"] = list(self.sheets.values())
        for key in coverage(0):
            if key not in {"entries", "capsHit"}:
                self.counts[key] = sum(scope[key] for scope in self.sheets.values())
        self.counts["capsHit"] = sorted(
            {cap for scope in self.sheets.values() for cap in scope["capsHit"]}
        )
