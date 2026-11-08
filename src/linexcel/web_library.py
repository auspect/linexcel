"""Library adapter: authorization and source mutation use the task-store lock."""

from __future__ import annotations

import sqlite3

from linexcel.catalog import Catalog, Source, fts5_available, revision


class Library:
    def __init__(self, store):
        self.store = store
        self.catalogs: dict[str, Catalog] = {}
        self.fts5 = fts5_available()

    def catalog(self, owner: str) -> Catalog:
        if owner not in self.catalogs:
            self.catalogs[owner] = Catalog(
                self.store.root / owner / "catalog", self.store.config.catalog_limits
            )
        return self.catalogs[owner]

    def sources(self, owner: str) -> list[Source]:
        from linexcel.web import _read

        sources = []
        for manifest in (self.store.root / owner / "projects").glob("*/project.json"):
            if (owner, manifest.parent.name) in self.store.unreadable_projects:
                continue
            project = _read(manifest)
            path = self.store.project_path(owner, project["id"])
            sources.append(
                Source(
                    project["id"],
                    project["filename"],
                    path / "index.sqlite",
                    project["createdAt"],
                )
            )
        return sources

    def refresh(self, owner: str):
        """Called under store.lock after import/delete, including replacement."""
        try:
            self.catalog(owner).ensure(self.sources(owner))
        except (OSError, ValueError, sqlite3.Error):
            # An optional derived index must not roll back a completed import.
            # The catalog persists a failed build state and the API reports it.
            pass

    def query(self, owner: str, operation: str, params: dict) -> dict:
        from linexcel.web import APIError

        allowed = {
            "search": {
                "q",
                "query",
                "kinds",
                "project",
                "revision",
                "offset",
                "limit",
                "generation",
            },
            "status": set(),
            "lineage": {"handle", "depth", "limit"},
        }
        if operation not in allowed or set(params) - allowed[operation]:
            raise APIError(400, "Invalid library parameter")

        def value(key, default=""):
            values = params.get(key, [default])
            if (
                len(values) != 1
                or not isinstance(values[0], str)
                or len(values[0]) > 4096
            ):
                raise ValueError("Invalid library parameter")
            return values[0]

        try:
            with self.store.lock:
                catalog = self.catalog(owner)
                sources = self.sources(owner)
                if operation == "lineage":
                    return catalog.lineage(
                        sources,
                        value("handle"),
                        depth=int(value("depth", "1")),
                        limit=int(value("limit", "200")),
                    )
                if operation == "status":
                    try:
                        catalog.ensure(sources)
                    except (OSError, ValueError, sqlite3.Error):
                        if catalog.build_state["status"] != "failed":
                            raise
                        return catalog.failed_status()
                    return catalog.status()
                if "q" in params and "query" in params:
                    raise ValueError("Supply q or query, not both")
                return catalog.search(
                    sources,
                    value("q", value("query")),
                    kinds=value("kinds").split(",") if value("kinds") else (),
                    project=value("project") or None,
                    revision_id=value("revision") or None,
                    offset=int(value("offset", "0")),
                    limit=int(value("limit", "50")),
                    generation=value("generation") or None,
                )
        except KeyError:
            raise APIError(404, "Library revision or cell not found") from None
        except (ValueError, OSError, sqlite3.Error) as error:
            state = self.catalog(owner).build_state
            if state["status"] == "failed":
                raise APIError(
                    503,
                    "Library build failed; inspect status and budgets",
                ) from None
            raise APIError(400, str(error)) from None


def project_revision(project: dict, path) -> dict:
    """Read identity from the immutable index; old projects need no migration."""
    from linexcel.lazy_store import WorkbookIndex, file_sha256

    index = path / "index.sqlite"
    sha = None
    if index.is_file():
        try:
            with WorkbookIndex(index) as store:
                sha = store.meta.get("sha256")
        except (OSError, sqlite3.Error, ValueError):
            pass  # Catalog status exposes the separate rebuild requirement.
    if not sha and (path / "workbook.xlsx").is_file():
        sha = file_sha256(path / "workbook.xlsx")
    return {
        **project,
        "workbookRevision": revision(project["id"], sha) if sha else None,
    }
