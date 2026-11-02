"""Private versioned workspace archives, bounded extraction and atomic restore."""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import shutil
import tempfile
import time
import uuid
import zipfile
from pathlib import Path, PurePosixPath

FORMAT = "linexcel-workspace"
VERSION = 1
CHUNK = 1024 * 1024
MANIFEST_LIMIT = 4 * CHUNK
MAX_FILES = 100000


class PrivateDownload:
    """Delete the temporary archive on completion or an abandoned download."""

    def __init__(self, path: Path):
        self.path = path
        try:
            self.stream = path.open("rb")
        except BaseException:
            path.unlink(missing_ok=True)
            raise

    def __iter__(self):
        return self

    def __next__(self):
        if self.stream.closed:
            raise StopIteration
        try:
            value = self.stream.read(65536)
            if not value:
                self.close()
                raise StopIteration
            return value
        except BaseException:
            self.close()
            raise

    def close(self):
        self.stream.close()
        self.path.unlink(missing_ok=True)


def _error(message: str, status: int = 400):
    from linexcel.web import APIError

    return APIError(status, message)


def _check_space(store, extra: int = 0) -> None:
    reserved = store.uploads.remaining_bytes()
    if store._storage_bytes() + reserved + extra > store.config.max_storage_mb * CHUNK:
        raise _error("Stockage insuffisant pour la sauvegarde ou restauration", 413)


def _valid_name(name: str) -> bool:
    if not isinstance(name, str) or len(name) > 240 or "\x00" in name:
        return False
    parts = PurePosixPath(name).parts
    if (
        "\\" in name
        or not parts
        or ".." in parts
        or name.startswith("/")
        or PurePosixPath(name).as_posix() != name
    ):
        return False
    if name == "workbook.xlsx":
        return True
    if len(parts) == 2 and parts[0] == "references":
        from linexcel.web import TaskStore

        try:
            return TaskStore._filename(parts[1]) == parts[1]
        except Exception:
            return False
    if len(parts) != 3 or not re.fullmatch(r"[a-f0-9]{32}", parts[1]):
        return False
    if parts[0] == "tasks":
        return parts[2] in {
            "result-data.json",
            "result-preview.json",
            "investigation.html",
        }
    return parts[0] == "captures" and bool(
        re.fullmatch(r"\d+(?:\.thumb)?\.png", parts[2])
    )


def backup(store, owner: str, project_id: str) -> Path:
    """Create an archive; caller must unlink the returned private temporary file."""
    from linexcel.web import TERMINAL, _read

    key = (owner, project_id)
    with store.lock:
        source = store.project_path(owner, project_id)
        tasks = [
            t.copy()
            for t in store.tasks.values()
            if (t["owner"], t["projectId"]) == key
        ]
        if any(t["status"] not in TERMINAL for t in tasks):
            raise _error("Attendez la fin des opérations avant la sauvegarde", 409)
        project = _read(source / "project.json")
        if store._revision(source) != project["revision"]:
            raise _error("Sources modifiées ; réimportez avant la sauvegarde", 409)
        store._project_exports[key] = store._project_exports.get(key, 0) + 1
        files = [source / "workbook.xlsx", *sorted((source / "references").glob("*"))]
        for task in tasks:
            task.pop("owner", None)
            task.pop("projectId", None)
            for name in (
                "result-data.json",
                "result-preview.json",
                "investigation.html",
            ):
                path = source / "tasks" / task["id"] / name
                if path.is_file():
                    files.append(path)
            files.extend(sorted((source / "captures" / task["id"]).glob("*.png")))
        directory = store.root / ".archives"
        directory.mkdir(exist_ok=True)
        descriptor, filename = tempfile.mkstemp(
            prefix="backup-", suffix=".zip", dir=directory
        )
        os.close(descriptor)
        destination = Path(filename)
    try:
        if len(files) > MAX_FILES:
            raise _error("Trop de fichiers dans cet espace", 413)
        manifest = {
            "format": FORMAT,
            "version": VERSION,
            "createdAt": time.time(),
            "project": {
                k: project[k]
                for k in ("name", "filename", "revision", "references", "createdAt")
                if k in project
            },
            "tasks": tasks,
            "files": [],
        }
        with zipfile.ZipFile(
            destination, "w", compression=zipfile.ZIP_DEFLATED
        ) as archive:
            for path in files:
                name = path.relative_to(source).as_posix()
                if not _valid_name(name) or path.is_symlink() or not path.is_file():
                    raise _error("Fichier de sauvegarde non pris en charge")
                digest, size = hashlib.sha256(), 0
                with (
                    path.open("rb") as stream,
                    archive.open(name, "w", force_zip64=True) as target,
                ):
                    while chunk := stream.read(CHUNK):
                        target.write(chunk)
                        digest.update(chunk)
                        size += len(chunk)
                        _check_space(store)
                manifest["files"].append(
                    {"path": name, "bytes": size, "sha256": digest.hexdigest()}
                )
            raw = json.dumps(manifest, ensure_ascii=False, allow_nan=False).encode()
            if len(raw) > MANIFEST_LIMIT:
                raise _error("Manifeste de sauvegarde trop volumineux", 413)
            archive.writestr("manifest.json", raw)
        _check_space(store)
        return destination
    except BaseException:
        destination.unlink(missing_ok=True)
        raise
    finally:
        with store.lock:
            count = store._project_exports[key] - 1
            if count:
                store._project_exports[key] = count
            else:
                store._project_exports.pop(key, None)


def receive_restore(store, owner: str, stream, length: int) -> dict:
    if not 0 < length <= store.config.max_upload_mb * CHUNK:
        raise _error("Archive trop volumineuse ou taille invalide", 413)
    directory = store.root / ".archives"
    directory.mkdir(exist_ok=True)
    descriptor, filename = tempfile.mkstemp(
        prefix="restore-", suffix=".zip", dir=directory
    )
    archive_path = Path(filename)
    try:
        with os.fdopen(descriptor, "wb") as target:
            remaining = length
            while remaining:
                chunk = stream.read(min(CHUNK, remaining))
                if not chunk:
                    raise _error("Transfert de sauvegarde incomplet")
                if len(chunk) > remaining:
                    raise _error("Taille du transfert incohérente")
                target.write(chunk)
                remaining -= len(chunk)
                _check_space(store)
        return restore(store, owner, archive_path)
    finally:
        archive_path.unlink(missing_ok=True)


def restore(store, owner: str, archive_path: Path) -> dict:
    from linexcel.web import MAX_RESULT_BYTES, OPERATIONS, TERMINAL, _write

    stage = None
    published_path = None
    published = False
    try:
        with zipfile.ZipFile(archive_path) as archive:
            members = archive.infolist()
            if len(members) > MAX_FILES + 1 or len(
                {m.filename for m in members}
            ) != len(members):
                raise _error("Archive contient trop de fichiers ou des doublons")
            info = archive.getinfo("manifest.json")
            if info.file_size > MANIFEST_LIMIT:
                raise _error("Manifeste de sauvegarde trop volumineux", 413)
            manifest = json.loads(archive.read(info))
            if (
                not isinstance(manifest, dict)
                or manifest.get("format") != FORMAT
                or manifest.get("version") != VERSION
            ):
                raise _error("Version de sauvegarde non prise en charge")
            project = manifest["project"]
            tasks = manifest["tasks"]
            entries = manifest["files"]
            if (
                not isinstance(project, dict)
                or not isinstance(tasks, list)
                or not isinstance(entries, list)
            ):
                raise _error("Manifeste de sauvegarde invalide")
            filename = store._filename(project.get("filename"))
            name = project.get("name")
            references = project.get("references", [])
            if (
                not isinstance(name, str)
                or len(name) > 200
                or not isinstance(references, list)
                or len(references) > 32
            ):
                raise _error("Métadonnées du classeur invalides")
            references = [store._filename(value) for value in references]
            if len(
                {filename.casefold(), *(value.casefold() for value in references)}
            ) != 1 + len(references):
                raise _error("Noms de fichiers en double")
            listed = {e["path"]: e for e in entries}
            if len(listed) != len(entries) or set(listed) | {"manifest.json"} != {
                m.filename for m in members
            }:
                raise _error("Inventaire de sauvegarde incohérent")
            if "workbook.xlsx" not in listed or {
                p for p in listed if p.startswith("references/")
            } != {"references/" + r for r in references}:
                raise _error("Fichiers sources manquants")
            total = 0
            for item in members:
                if item.filename == "manifest.json":
                    continue
                entry = listed[item.filename]
                if (
                    not _valid_name(item.filename)
                    or item.is_dir()
                    or item.flag_bits & 1
                ):
                    raise _error("Chemin de sauvegarde invalide")
                if entry.get("bytes") != item.file_size or not re.fullmatch(
                    r"[a-f0-9]{64}", str(entry.get("sha256", ""))
                ):
                    raise _error("Empreinte ou taille de sauvegarde invalide")
                if (
                    item.filename.endswith(("result-data.json", "result-preview.json"))
                    and item.file_size > MAX_RESULT_BYTES
                ):
                    raise _error("Résultat sauvegardé trop volumineux", 413)
                total += item.file_size
            task_ids = {}
            for task in tasks:
                if (
                    not isinstance(task, dict)
                    or not re.fullmatch(r"[a-f0-9]{32}", str(task.get("id", "")))
                    or task.get("operation") not in OPERATIONS
                    or task.get("status") not in TERMINAL
                ):
                    raise _error("Opération sauvegardée invalide")
                if (
                    not isinstance(task.get("createdAt"), (int, float))
                    or isinstance(task.get("createdAt"), bool)
                    or not math.isfinite(task["createdAt"])
                    or not isinstance(task.get("options", {}), dict)
                    or (
                        task.get("result") is not None
                        and not isinstance(task["result"], dict)
                    )
                ):
                    raise _error("Métadonnées d’opération invalides")
                if task["id"] in task_ids:
                    raise _error("Opération sauvegardée en double")
                task_ids[task["id"]] = uuid.uuid4().hex
            for path in listed:
                parts = PurePosixPath(path).parts
                if parts[0] in {"tasks", "captures"} and parts[1] not in task_ids:
                    raise _error("Fichier sans opération correspondante")
            with store.lock:
                store._has_capacity()
                _check_space(store, total + len(tasks) * 16384 + 65536)
                stage = Path(
                    tempfile.mkdtemp(prefix="restore-stage-", dir=directory_for(store))
                )
                (stage / "references").mkdir()
                for path, entry in listed.items():
                    parts = list(PurePosixPath(path).parts)
                    if parts[0] in {"tasks", "captures"}:
                        parts[1] = task_ids[parts[1]]
                    destination = stage.joinpath(*parts)
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    digest = hashlib.sha256()
                    with archive.open(path) as source, destination.open("wb") as target:
                        copied = 0
                        while chunk := source.read(CHUNK):
                            copied += len(chunk)
                            if copied > entry["bytes"]:
                                raise _error("Taille décompressée incohérente")
                            digest.update(chunk)
                            target.write(chunk)
                    if (
                        copied != entry["bytes"]
                        or digest.hexdigest() != entry["sha256"]
                    ):
                        raise _error(
                            "Empreinte de fichier incorrecte ; restauration annulée"
                        )
                    if path.endswith(("result-data.json", "result-preview.json")):
                        value = json.loads(destination.read_text("utf-8"))
                        if not isinstance(value, dict):
                            raise _error("Résultat sauvegardé invalide")
                from linexcel.macros import valid_excel_header

                for logical, path in [
                    (filename, stage / "workbook.xlsx"),
                    *((r, stage / "references" / r) for r in references),
                ]:
                    with path.open("rb") as stream:
                        if not valid_excel_header(stream.read(8), logical):
                            raise _error("Fichier source invalide dans la sauvegarde")
                project_id = uuid.uuid4().hex
                restored = {
                    "id": project_id,
                    "name": name,
                    "filename": filename,
                    "references": references,
                    "createdAt": time.time(),
                    "revision": store._revision(stage),
                    "restoredAt": time.time(),
                    "originalCreatedAt": project.get("createdAt"),
                }
                if restored["revision"] != project.get("revision"):
                    raise _error("Révision des sources différente du manifeste")
                _write(stage / "project.json", restored)
                restored_tasks = []
                for original in tasks:
                    task = dict(original)
                    old_id, new_id = task["id"], task_ids[task["id"]]
                    task.update(
                        id=new_id,
                        owner=owner,
                        projectId=project_id,
                        restoredFromTask=old_id,
                        historical=True,
                    )
                    task["cacheVersion"] = "restored-historical"
                    task.setdefault("createdAt", time.time())
                    task.setdefault("nodeId", None)
                    task.setdefault("result", None)
                    task["revision"] = restored["revision"]
                    options = dict(task.get("options") or {})
                    if options.get("captureTaskId") in task_ids:
                        options["captureTaskId"] = task_ids[options["captureTaskId"]]
                    task["options"] = options
                    root = stage / "tasks" / new_id
                    root.mkdir(parents=True, exist_ok=True)
                    if task.get("resultAvailable") and not all(
                        (root / f).is_file()
                        for f in ("result-data.json", "result-preview.json")
                    ):
                        raise _error("Résultat sauvegardé manquant")
                    if task["operation"] == "capture" and task.get("resultAvailable"):
                        for part in ("result-data.json", "result-preview.json"):
                            if (root / part).stat().st_size > MANIFEST_LIMIT:
                                raise _error("Métadonnées de capture trop volumineuses")
                            value = json.loads((root / part).read_text("utf-8"))
                            for index, shot in enumerate(value.get("screenshots", [])):
                                shot["url"] = (
                                    f"api/projects/{project_id}/captures/{new_id}/{index}"
                                )
                                if "thumbnailUrl" in shot:
                                    shot["thumbnailUrl"] = shot["url"] + "?thumbnail=1"
                            _write(root / part, value)
                    _write(root / "task.json", task)
                    restored_tasks.append(task)
                destination = store.root / owner / "projects" / project_id
                destination.parent.mkdir(parents=True, exist_ok=True)
                stage.rename(destination)
                published_path = destination
                stage = None
                for task in restored_tasks:
                    store.tasks[task["id"]] = task
                imported = store.submit(
                    owner, project_id, "import", None, {}, force=True
                )
                published = True
                return {"project": restored, "task": imported}
    except (zipfile.BadZipFile, KeyError, TypeError, ValueError, UnicodeError) as exc:
        raise _error("Sauvegarde invalide ou incomplète") from exc
    finally:
        if stage is not None:
            shutil.rmtree(stage, ignore_errors=True)
        if not published and published_path is not None:
            with store.lock:
                for task_id, task in list(store.tasks.items()):
                    if (
                        task["owner"] == owner
                        and task["projectId"] == published_path.name
                    ):
                        store.tasks.pop(task_id, None)
                        store.events.pop(task_id, None)
                shutil.rmtree(published_path, ignore_errors=True)


def directory_for(store) -> Path:
    directory = store.root / ".archives"
    directory.mkdir(exist_ok=True)
    return directory


def clean_interrupted_archives(root: Path) -> None:
    """A coordinator owns the directory exclusively before this startup cleanup."""
    directory = root / ".archives"
    if not directory.is_dir() or directory.is_symlink():
        return
    for path in directory.iterdir():
        if path.is_symlink():
            continue
        if path.is_file() and re.fullmatch(
            r"(?:backup|restore)-[a-z0-9_]+\.zip", path.name
        ):
            path.unlink(missing_ok=True)
        elif path.is_dir() and re.fullmatch(r"restore-stage-[a-z0-9_]+", path.name):
            shutil.rmtree(path)


def legacy_projects(store, owner: str) -> list[dict]:
    from linexcel.web import _read

    result = []
    for path in store.root.glob("*/projects/*/project.json"):
        old_owner = path.parent.parent.parent.name
        if old_owner == owner or not re.fullmatch(r"[a-f0-9]{64}", old_owner):
            continue
        try:
            value = _read(path)
            if value.get("id") != path.parent.name or value.get("recoveredTo"):
                continue
            result.append(
                {
                    "id": hashlib.sha256(
                        (old_owner + "/" + value["id"]).encode()
                    ).hexdigest(),
                    "name": value["name"],
                    "createdAt": value["createdAt"],
                }
            )
        except (OSError, ValueError, KeyError, TypeError):
            continue
    return result


def recover_legacy(store, owner: str, identifier: str) -> dict:
    from linexcel.web import _read, _write

    with store.lock:
        for path in store.root.glob("*/projects/*/project.json"):
            old_owner, project_id = path.parent.parent.parent.name, path.parent.name
            digest = hashlib.sha256((old_owner + "/" + project_id).encode()).hexdigest()
            if (
                old_owner == owner
                or digest != identifier
                or not re.fullmatch(r"[a-f0-9]{64}", old_owner)
            ):
                continue
            previous = _read(path)
            if previous.get("recoveredTo"):
                raise _error("Cet espace a déjà été récupéré", 409)
            archive = backup(store, old_owner, project_id)
            try:
                result = restore(store, owner, archive)
                project = _read(path)
                project["recoveredTo"] = result["project"]["id"]
                _write(path, project)
                return result
            finally:
                archive.unlink(missing_ok=True)
    raise _error("Ancien espace introuvable", 404)
