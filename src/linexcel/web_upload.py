"""Bounded, owner-scoped uploads; workbook bytes are never JSON payloads.

Each acknowledged chunk is committed to disk before its offset is published.
An interrupted chunk is discarded, so clients can safely resume at ``received``.
"""

from __future__ import annotations

import os
import shutil
import stat
import threading
import time
import uuid
from contextlib import ExitStack
from pathlib import Path
from typing import Any, BinaryIO

CHUNK_BYTES = 8 * 1024 * 1024
COPY_BYTES = 1024 * 1024
UPLOAD_LIFETIME = 24 * 3600


class UploadStore:
    def __init__(self, store):
        self.store = store
        self.sessions: dict[str, dict] = {}
        self.completed: dict[str, dict] = {}
        self.locks: dict[str, threading.Lock] = {}
        self._recover()

    def _recover(self):
        from linexcel.web import _read

        for directory in self.store.root.glob("*/uploads/*"):
            path = directory / "upload.json"
            try:
                session = _read(path)
                if (
                    session["id"] != directory.name
                    or session["owner"] != directory.parent.parent.name
                ):
                    raise ValueError("Invalid upload owner")
                if session["updatedAt"] < time.time() - UPLOAD_LIFETIME:
                    raise ValueError("Expired upload")
                for index, item in enumerate(session["files"]):
                    file = self._file_path(path.parent, session, index)
                    received = item["received"]
                    if received and (
                        not file.exists() or file.stat().st_size < received
                    ):
                        raise ValueError("Incomplete upload")
                    if file.exists():
                        with file.open("r+b") as stream:
                            stream.truncate(received)
                self.sessions[session["id"]] = session
                self.locks[session["id"]] = threading.Lock()
            except (OSError, ValueError, KeyError, TypeError):
                self._remove_path(directory)
        for path in self.store.root.glob("*/projects/*/project.json"):
            project = _read(path)
            if not project.get("uploadId"):
                continue
            task = next(
                (
                    task
                    for task in self.store.tasks.values()
                    if (
                        task["projectId"] == project["id"]
                        and task["operation"] == "import"
                    )
                ),
                None,
            )
            if task:
                self.completed[project["uploadId"]] = {
                    "owner": task["owner"],
                    "project": project,
                    "task": self.store.public(task),
                }

    def _path(self, session: dict) -> Path:
        return self.store.root / session["owner"] / "uploads" / session["id"]

    @staticmethod
    def _file_path(path: Path, session: dict, index: int) -> Path:
        return (
            path / "workbook.xlsx"
            if index == 0
            else path / "references" / session["files"][index]["name"]
        )

    @staticmethod
    def public(session: dict) -> dict:
        return {
            "id": session["id"],
            "files": [dict(item) for item in session["files"]],
            "chunkBytes": CHUNK_BYTES,
        }

    def remaining_bytes(self) -> int:
        return sum(
            item["size"] - item["received"]
            for session in self.sessions.values()
            for item in session["files"]
        )

    def _get(self, owner: str, upload_id: str) -> dict:
        from linexcel.web import APIError

        session = self.sessions.get(upload_id)
        if session is None or session["owner"] != owner:
            raise APIError(404, "Transfert introuvable ou expiré")
        return session

    def get(self, owner: str, upload_id: str) -> dict:
        with self.store.lock:
            return {"upload": self.public(self._get(owner, upload_id))}

    def _expire(self):
        for upload_id, session in list(self.sessions.items()):
            if session["updatedAt"] >= time.time() - UPLOAD_LIFETIME:
                continue
            lock = self.locks[upload_id]
            if lock.acquire(blocking=False):
                try:
                    self._remove(session)
                finally:
                    lock.release()

    def _remove(self, session: dict):
        self._remove_path(self._path(session))
        self.sessions.pop(session["id"], None)
        self.locks.pop(session["id"], None)

    def _remove_path(self, path: Path):
        relative = path.resolve().relative_to(self.store.root.resolve())
        if len(relative.parts) != 3 or relative.parts[1] != "uploads":
            raise ValueError("Unsafe upload cleanup path")
        shutil.rmtree(path, ignore_errors=True)

    def begin(self, owner: str, body: dict) -> dict:
        from linexcel.web import APIError, _write

        budget = self.store.budget(body.get("budget", {}))
        refs = body.get("references", [])
        if not isinstance(refs, list) or len(refs) > 32:
            raise APIError(400, "Au plus 32 classeurs de référence")
        files: list[dict[str, Any]] = []
        for item in [body.get("workbook"), *refs]:
            if not isinstance(item, dict):
                raise APIError(400, "Fichier invalide")
            name = self.store._filename(item.get("name"))
            size = item.get("size")
            if isinstance(size, bool) or not isinstance(size, int) or size < 2:
                raise APIError(400, "Taille de fichier invalide")
            files.append({"name": name, "size": size, "received": 0})
        if len({item["name"].casefold() for item in files}) != len(files):
            raise APIError(400, "Les noms des fichiers doivent être uniques")
        total = sum(item["size"] for item in files)
        if total > self.store.config.max_upload_mb * 1048576:
            raise APIError(413, "Import trop volumineux")
        name = body.get("name") or files[0]["name"]
        if not isinstance(name, str) or len(name) > 200:
            raise APIError(400, "Nom de l’espace invalide")
        with self.store.lock:
            self._expire()
            self.store._has_capacity()
            if len(self.sessions) >= self.store.config.max_pending_tasks:
                raise APIError(429, "Trop de transferts en attente")
            # Reserve all announced bytes, including concurrent incomplete uploads.
            if (
                self.store._storage_bytes() + self.remaining_bytes() + total + 65536
                > self.store.config.max_storage_mb * 1048576
            ):
                raise APIError(413, "Stockage plein ; supprimez des classeurs")
            session = {
                "id": uuid.uuid4().hex,
                "owner": owner,
                "name": name,
                "files": files,
                "budget": budget,
                "createdAt": time.time(),
                "updatedAt": time.time(),
            }
            path = self._path(session)
            try:
                (path / "references").mkdir(parents=True)
                _write(path / "upload.json", session)
            except OSError:
                self._remove_path(path)
                raise APIError(
                    507, "Impossible de réserver le stockage du transfert"
                ) from None
            self.sessions[session["id"]] = session
            self.locks[session["id"]] = threading.Lock()
            return {"upload": self.public(session)}

    def _acquire(self, owner: str, upload_id: str):
        from linexcel.web import APIError

        with self.store.lock:
            session = self._get(owner, upload_id)
            lock = self.locks[upload_id]
            if not lock.acquire(blocking=False):
                raise APIError(409, "Un transfert est déjà en cours pour ce fichier")
            return session, lock

    def _release(self, session: dict, lock):
        with self.store.lock:
            if session.get("cancelRequested"):
                self._remove(session)
            lock.release()

    def append(
        self,
        owner: str,
        upload_id: str,
        index: int,
        offset: int,
        length: int,
        stream: BinaryIO,
    ) -> dict:
        from linexcel.web import APIError, _write

        if not 0 < length <= CHUNK_BYTES:
            raise APIError(413, "Fragment limité à 8 Mio")
        session, lock = self._acquire(owner, upload_id)
        try:
            if not 0 <= index < len(session["files"]):
                raise APIError(404, "Fichier du transfert introuvable")
            item = session["files"][index]
            if offset != item["received"]:
                raise APIError(
                    409, "Position incorrecte ; reprenez le transfert enregistré"
                )
            if offset + length > item["size"]:
                raise APIError(413, "Fragment au-delà de la taille annoncée")
            with self.store.lock:
                if self.store.closed:
                    raise APIError(503, "Serveur en cours d’arrêt")
                if (
                    self.store._storage_bytes() + self.remaining_bytes()
                    > self.store.config.max_storage_mb * 1048576
                ):
                    raise APIError(413, "Stockage plein ; supprimez des classeurs")
            path = self._file_path(self._path(session), session, index)
            try:
                with path.open("r+b" if path.exists() else "w+b") as target:
                    target.seek(offset)
                    remaining = length
                    try:
                        while remaining:
                            if session.get("cancelRequested"):
                                raise APIError(409, "Transfert annulé")
                            try:
                                chunk = stream.read(min(COPY_BYTES, remaining))
                            except OSError:
                                raise APIError(
                                    400, "Connexion interrompue ; reprenez le transfert"
                                ) from None
                            if not chunk:
                                raise APIError(
                                    400, "Transfert interrompu ; fragment incomplet"
                                )
                            if len(chunk) > remaining:
                                raise APIError(400, "Fragment invalide")
                            target.write(chunk)
                            remaining -= len(chunk)
                        if offset == 0 and length >= 2:
                            target.seek(0)
                            if target.read(2) != b"PK":
                                raise APIError(
                                    400, "Le fichier doit être un classeur OOXML"
                                )
                        target.flush()
                        os.fsync(target.fileno())
                        with self.store.lock:
                            if session.get("cancelRequested"):
                                raise APIError(409, "Transfert annulé")
                            updated = {
                                **session,
                                "files": [dict(f) for f in session["files"]],
                            }
                            updated["files"][index]["received"] = offset + length
                            updated["updatedAt"] = time.time()
                            _write(self._path(session) / "upload.json", updated)
                            session.update(updated)
                    except BaseException:
                        target.truncate(offset)
                        raise
            except OSError:
                raise APIError(
                    507, "Écriture interrompue ; vérifiez le stockage disponible"
                ) from None
            return {"upload": self.public(session)}
        finally:
            self._release(session, lock)

    def cancel(self, owner: str, upload_id: str) -> dict:
        with self.store.lock:
            session = self._get(owner, upload_id)
            lock = self.locks[upload_id]
            if not lock.acquire(blocking=False):
                # Abort remains immediate even while a socket is blocked in read.
                # The owning request rolls back and removes the staging files.
                session["cancelRequested"] = True
                return {}
            try:
                self._remove(session)
                return {}
            finally:
                lock.release()

    def complete(self, owner: str, upload_id: str) -> dict:
        from linexcel.web import APIError, _write

        with self.store.lock:
            completed = self.completed.get(upload_id)
            if completed and completed["owner"] == owner:
                self.store.project_path(owner, completed["project"]["id"])
                return {"project": completed["project"], "task": completed["task"]}
        session, lock = self._acquire(owner, upload_id)
        try:
            path = self._path(session)
            for index, item in enumerate(session["files"]):
                if item["received"] != item["size"]:
                    raise APIError(409, "Transfert incomplet")
                file = self._file_path(path, session, index)
                if file.stat().st_size != item["size"]:
                    raise APIError(409, "Fichier du transfert incomplet")
                with file.open("rb") as source:
                    if source.read(2) != b"PK":
                        raise APIError(400, "Le fichier doit être un classeur OOXML")
            revision = self.store._revision(path)
            with self.store.lock:
                if session.get("cancelRequested"):
                    raise APIError(409, "Transfert annulé")
                self.store._has_capacity()
                project_id = uuid.uuid4().hex
                destination = self.store.root / owner / "projects" / project_id
                destination.parent.mkdir(parents=True, exist_ok=True)
                project = {
                    "id": project_id,
                    "uploadId": upload_id,
                    "name": session["name"],
                    "filename": session["files"][0]["name"],
                    "revision": revision,
                    "references": [item["name"] for item in session["files"][1:]],
                    "createdAt": time.time(),
                }
                _write(path / "project.json", project)
                path.rename(destination)
                try:
                    task = self.store.submit(
                        owner, project_id, "import", None, session["budget"]
                    )
                except BaseException:
                    destination.rename(path)
                    (path / "project.json").unlink(missing_ok=True)
                    raise
                self.sessions.pop(upload_id)
                self.locks.pop(upload_id)
                self.completed[upload_id] = {
                    "owner": owner,
                    "project": project,
                    "task": task,
                }
                # Keep the harmless manifest if a cleanup failure occurs after
                # the project/task commit; success must remain unambiguous.
                try:
                    (destination / "upload.json").unlink(missing_ok=True)
                except OSError:
                    pass
                return {"project": project, "task": task}
        finally:
            self._release(session, lock)

    def local(self, owner: str, body: dict) -> dict:
        """Copy a local snapshot, never retain mutable paths to the user's source."""
        from linexcel.web import APIError

        references = body.get("references", [])
        if not isinstance(references, list) or len(references) > 32:
            raise APIError(400, "Au plus 32 classeurs de référence")
        with ExitStack() as opened:
            sources = []
            for value in [body.get("path"), *references]:
                if not isinstance(value, str) or not value or "\x00" in value:
                    raise APIError(400, "Chemin local absolu requis")
                path = Path(value)
                if not path.is_absolute() or value.startswith(("\\\\", "//")):
                    raise APIError(400, "Chemin de fichier local absolu requis")
                self.store._filename(path.name)
                try:
                    source = opened.enter_context(path.open("rb"))
                    details = os.fstat(source.fileno())
                    if not stat.S_ISREG(details.st_mode):
                        raise OSError("Not a regular file")
                except OSError:
                    raise APIError(400, "Fichier local inaccessible") from None
                sources.append((path.name, source, details))
            manifest = [
                {"name": name, "size": info.st_size} for name, _, info in sources
            ]
            response = self.begin(
                owner,
                {
                    "workbook": manifest[0],
                    "references": manifest[1:],
                    "name": body.get("name"),
                    "budget": body.get("budget", {}),
                },
            )
            upload_id = response["upload"]["id"]
            try:
                for index, (_, source, initial) in enumerate(sources):
                    offset = 0
                    while offset < initial.st_size:
                        length = min(CHUNK_BYTES, initial.st_size - offset)
                        self.append(owner, upload_id, index, offset, length, source)
                        offset += length
                    current = os.fstat(source.fileno())
                    if (current.st_size, current.st_mtime_ns) != (
                        initial.st_size,
                        initial.st_mtime_ns,
                    ):
                        raise APIError(
                            409, "Source modifiée pendant la copie ; recommencez"
                        )
                return self.complete(owner, upload_id)
            except BaseException:
                self.cancel(owner, upload_id)
                raise
