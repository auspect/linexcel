"""Complete project JSON exports with bounded graph and result buffers."""

from __future__ import annotations

import json
from collections.abc import Generator, Iterator

from linexcel.lazy_store import WorkbookIndex, _pattern

CHUNK_BYTES = 65536
_ENCODER = json.JSONEncoder(ensure_ascii=False, allow_nan=False, separators=(",", ":"))


def _json(value) -> Iterator[bytes]:
    for token in _ENCODER.iterencode(value):
        # Split characters first so a single long source string cannot create a
        # second workbook-sized byte buffer. UTF-8 uses at most four bytes/char.
        for start in range(0, len(token), CHUNK_BYTES // 4):
            yield token[start : start + CHUNK_BYTES // 4].encode("utf-8")


def _file(path) -> Iterator[bytes]:
    with path.open("rb") as source:
        while chunk := source.read(CHUNK_BYTES):
            yield chunk


def _buffer(fragments) -> Generator[bytes, None, None]:
    pending = bytearray()
    try:
        for fragment in fragments:
            pending.extend(fragment)
            while len(pending) >= CHUNK_BYTES:
                yield bytes(pending[:CHUNK_BYTES])
                del pending[:CHUNK_BYTES]
        if pending:
            yield bytes(pending)
    finally:
        fragments.close()


class ProjectExport:
    """Own the export lease even if WSGI closes before the first yielded byte."""

    def __init__(self, store, owner: str, project_id: str):
        from linexcel.web import TERMINAL, APIError, _read

        self.store = store
        self.key = (owner, project_id)
        self.index = None
        self.closed = False
        self._iterator = None
        self._leased = False
        try:
            with store.lock:
                self.path = store.project_path(owner, project_id)
                tasks = [
                    t
                    for t in store.tasks.values()
                    if (t["owner"], t["projectId"]) == self.key
                ]
                if any(
                    t["operation"] == "import" and t["status"] not in TERMINAL
                    for t in tasks
                ):
                    raise APIError(409, "Attendez la fin de l’import avant l’export")
                store._project_exports[self.key] = (
                    store._project_exports.get(self.key, 0) + 1
                )
                self._leased = True
                self.project = _read(self.path / "project.json")
                self.tasks = [store.public(t, include_result=False) for t in tasks]
                index_path = self.path / "index.sqlite"
                if index_path.is_file():
                    self.index = WorkbookIndex(index_path)
                    self.index.db.execute("BEGIN")
                    self.meta = self.index.meta
                    self.count = len(self.index)
                self._iterator = _buffer(self._generate())
        except BaseException:
            self.close()
            raise

    def __iter__(self):
        return self

    def __next__(self):
        if self.closed or self._iterator is None:
            raise StopIteration
        try:
            return next(self._iterator)
        except BaseException:
            self.close()
            raise

    def close(self):
        if self.closed:
            return
        self.closed = True
        try:
            if self._iterator is not None:
                self._iterator.close()
        finally:
            try:
                if self.index is not None:
                    self.index.close()
            finally:
                if self._leased:
                    with self.store.lock:
                        remaining = self.store._project_exports[self.key] - 1
                        if remaining:
                            self.store._project_exports[self.key] = remaining
                        else:
                            self.store._project_exports.pop(self.key)
                    self._leased = False

    def _graph(self):
        if self.index is None:
            legacy = self.path / "graph.json"
            if legacy.is_file():
                yield from _file(legacy)
            else:
                yield b"null"
            return
        store = self.index
        yield b'{"lazy":true,"nodes":['
        first = True
        for row in store.db.execute("SELECT data,pattern_id FROM nodes ORDER BY seq"):
            if not first:
                yield b","
            first = False
            yield from _json(store._decode(row, context=True))
        yield b'],"edges":['
        first = True
        for source, target in store.db.execute(
            "SELECT source,target FROM edges ORDER BY source,target"
        ):
            if not first:
                yield b","
            first = False
            yield from _json({"source": source, "target": target, "kind": "dep"})
        yield b'],"formulaPatterns":['
        first = True
        for row in store.db.execute(
            "SELECT data,member_count FROM patterns ORDER BY sheet,id"
        ):
            if not first:
                yield b","
            first = False
            # Membership is recoverable from every exported node's patternId.
            # Keep convenience memberIds/ranges explicitly sampled and bounded.
            yield from _json(_pattern(store, row, member_limit=20))
        yield b'],"meta":'
        yield from _json(self.meta)
        yield b',"pagination":'
        yield from _json(
            {"offset": 0, "limit": self.count, "total": self.count, "hasMore": False}
        )
        yield b"}"

    def _generate(self):
        yield b'{"project":'
        yield from _json(self.project)
        yield b',"export":{"scope":"complete","patternMembers":"sampled; '
        yield b'complete membership is retained on nodes via patternId"},"graph":'
        yield from self._graph()
        yield b',"tasks":['
        for position, task in enumerate(self.tasks):
            if position:
                yield b","
            yield b"{"
            for key, value in task.items():
                if key == "result":
                    continue
                yield from _json(key)
                yield b":"
                yield from _json(value)
                yield b","
            yield b'"result":'
            result = self.path / "tasks" / task["id"] / "result-data.json"
            if task.get("resultAvailable"):
                yield from _file(result)
            else:
                yield from _json(task.get("result"))
            yield b"}"
        yield b"]}"
