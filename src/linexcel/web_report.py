"""Frozen, bounded investigations readable offline, including without JavaScript."""

from __future__ import annotations

import base64
import html
import json
import re
import time
from datetime import UTC, datetime
from itertools import islice
from pathlib import Path

from linexcel.lazy_store import WorkbookIndex, file_sha256, graph_neighborhood

MAX_TARGETS = 20
MAX_CAPTURES = 20
MAX_CAPTURE_BYTES = 64 * 1024 * 1024
MAX_SNAPSHOT_BYTES = 16 * 1024 * 1024
NODE_FIELDS = (
    "id",
    "kind",
    "sheet",
    "cell",
    "sourceFormula",
    "canonicalFormula",
    "formula",
    "value",
    "cachedValue",
    "cachedValueKind",
    "cache_status",
    "valueSource",
    "sourceDataType",
    "formulaProvenance",
    "sourceValueProvenance",
    "diagnostics",
    "reference",
    "bounds",
    "cell_count",
)


def normalize_options(options: dict) -> dict:
    from linexcel.web import APIError

    if not isinstance(options, dict) or set(options) - {
        "nodeIds",
        "captureIds",
        "language",
    }:
        raise APIError(400, "Options du rapport invalides")
    nodes = options.get("nodeIds")
    captures = options.get("captureIds", [])
    language = options.get("language", "fr")
    if (
        not isinstance(nodes, list)
        or not 1 <= len(nodes) <= MAX_TARGETS
        or any(not isinstance(n, str) or not n or len(n) > 1024 for n in nodes)
        or len(set(nodes)) != len(nodes)
    ):
        raise APIError(400, "Sélectionnez de 1 à 20 cellules distinctes")
    if (
        not isinstance(captures, list)
        or len(captures) > MAX_CAPTURES
        or any(
            not isinstance(c, str) or not re.fullmatch(r"[a-f0-9]{32}:[0-9]{1,6}", c)
            for c in captures
        )
        or len(set(captures)) != len(captures)
    ):
        raise APIError(400, "Sélection de captures invalide (20 au maximum)")
    if language not in ("fr", "en"):
        raise APIError(400, "Langue du rapport : fr ou en")
    return {"nodeIds": list(nodes), "captureIds": list(captures), "language": language}


def _node(node: dict) -> dict:
    return {key: node[key] for key in NODE_FIELDS if key in node}


def prepare_report(store, owner: str, project_id: str, options: dict) -> dict:
    """Called under the coordinator lock; freeze bounded evidence, never images."""
    from linexcel.web import APIError, _read

    options = normalize_options(options)
    project = store.project_path(owner, project_id)
    metadata = _read(project / "project.json")
    tasks = [
        t
        for t in store.tasks.values()
        if t["owner"] == owner and t["projectId"] == project_id
    ]
    if any(
        t["operation"] == "import"
        and t["status"] not in {"succeeded", "failed", "cancelled"}
        for t in tasks
    ):
        raise APIError(409, "Attendez la fin de l’import avant le rapport")
    if not (project / "index.sqlite").is_file():
        raise APIError(409, "Actualisez la structure avant le rapport")
    completed = sorted(
        (
            t
            for t in tasks
            if t["status"] == "succeeded" and t["revision"] == metadata["revision"]
        ),
        key=lambda t: (t["createdAt"], t["id"]),
        reverse=True,
    )

    def latest(operation, node_id):
        task = next(
            (
                t
                for t in completed
                if t["operation"] == operation
                and t.get("nodeId") == node_id
                and t.get("cacheVersion") == store.cache_version
                and (
                    operation != "document"
                    or t.get("options", {}).get("language") == options["language"]
                )
            ),
            None,
        )
        if task is None:
            return None
        result = store._task_result(task)
        if result is None:
            return None
        return {"id": task["id"], "createdAt": task["createdAt"], "result": result}

    def latest_attempt(operation, node_id):
        candidates = [
            t
            for t in tasks
            if t["operation"] == operation
            and t.get("nodeId") == node_id
            and t.get("revision") == metadata["revision"]
            and t.get("cacheVersion") == store.cache_version
            and (
                operation != "document"
                or t.get("options", {}).get("language") == options["language"]
            )
        ]
        if not candidates:
            return None
        task = max(candidates, key=lambda t: (t["createdAt"], t["id"]))
        return {key: task[key] for key in ("id", "createdAt", "status")}

    records = []
    evidence_bytes = 0
    with WorkbookIndex(project / "index.sqlite") as index:
        meta = index.meta
        for node_id in options["nodeIds"]:
            if node_id not in index:
                raise APIError(404, "Cellule sélectionnée introuvable")
            neighborhood = graph_neighborhood(project / "index.sqlite", node_id)
            nodes = {n["id"]: n for n in neighborhood["nodes"]}
            counts = {}
            for name, column in (("sources", "target"), ("uses", "source")):
                counts[name] = index.db.execute(
                    f"SELECT count(*) FROM edges WHERE {column}=?", (node_id,)
                ).fetchone()[0]
            ranges = []
            remaining = 200
            for node in nodes.values():
                if node.get("kind") != "range" or not node.get("bounds"):
                    continue
                r1, c1, r2, c2 = node["bounds"]
                count = index.db.execute(
                    "SELECT count(*) FROM nodes WHERE sheet=? "
                    "AND row BETWEEN ? AND ? AND col BETWEEN ? AND ? "
                    "AND kind IN ('cell','input')",
                    (node["sheet"], r1, r2, c1, c2),
                ).fetchone()[0]
                members = [
                    _node(index[member])
                    for member in islice(
                        index.range_members(node["sheet"], node["bounds"]), remaining
                    )
                ]
                remaining -= len(members)
                ranges.append(
                    {
                        "id": node["id"],
                        "representedCellCount": count,
                        "positionCount": (r2 - r1 + 1) * (c2 - c1 + 1),
                        "members": members,
                        "omitted": count - len(members),
                    }
                )
            records.append(
                {
                    "node": _node(nodes[node_id]),
                    "context": [_node(n) for n in nodes.values() if n["id"] != node_id],
                    "edges": neighborhood["edges"],
                    "relationCounts": counts,
                    "bounds": neighborhood["pagination"],
                    "depth": 1,
                    "ranges": ranges,
                    "calculation": latest("evaluate", node_id),
                    "documentation": latest("document", node_id),
                    "calculationAttempt": latest_attempt("evaluate", node_id),
                    "documentationAttempt": latest_attempt("document", node_id),
                }
            )
            evidence_bytes += len(json.dumps(records[-1], ensure_ascii=False).encode())
            if evidence_bytes > MAX_SNAPSHOT_BYTES:
                raise APIError(413, "Preuves trop volumineuses ; réduisez la sélection")
    captures = []
    total_bytes = 0
    for capture_id in options["captureIds"]:
        task_id, raw_index = capture_id.split(":")
        capture_index = int(raw_index)
        task = next(
            (
                t
                for t in completed
                if t["id"] == task_id and t["operation"] == "capture"
            ),
            None,
        )
        if task is None:
            raise APIError(404, "Capture de cette révision introuvable")
        path = store.capture_path(owner, project_id, task_id, capture_index)
        total_bytes += path.stat().st_size
        if total_bytes > MAX_CAPTURE_BYTES:
            raise APIError(413, "Captures du rapport limitées à 64 Mio")
        result = store._task_result(task) or {}
        shot = result["screenshots"][capture_index]
        captures.append(
            {
                "id": capture_id,
                "sheet": shot.get("sheet", ""),
                "name": shot.get("name", ""),
                "createdAt": task["createdAt"],
                "path": path.relative_to(project).as_posix(),
                "sha256": file_sha256(path),
            }
        )
    snapshot = {
        "reportVersion": 1,
        "createdAt": time.time(),
        "scope": options,
        "project": {
            key: metadata[key] for key in ("id", "name", "filename", "revision")
        },
        "sourceSha256": meta.get("sha256"),
        "nodeCount": meta.get("nodeCount"),
        "limitations": meta.get("limits", []),
        "records": records,
        "captures": captures,
    }
    if len(json.dumps(snapshot, ensure_ascii=False).encode()) > MAX_SNAPSHOT_BYTES:
        raise APIError(
            413, "Preuves du rapport trop volumineuses ; réduisez la sélection"
        )
    return snapshot


_CSS = (Path(__file__).parent / "assets" / "report.css").read_text("utf-8")


def _escape(value) -> str:
    return html.escape(str(value), quote=True)


def _value(value) -> str:
    return _escape(json.dumps(value, ensure_ascii=False, allow_nan=False))


def _when(timestamp) -> str:
    return datetime.fromtimestamp(timestamp, UTC).isoformat(timespec="seconds")


def _prose(markdown: str) -> str:
    """Small safe Markdown subset; all workbook/provider HTML remains text."""
    from linexcel.ai_validation import strip_claims

    lines = strip_claims(markdown).splitlines()
    output = []
    fenced = False
    table = False
    for line in lines:
        if line.lstrip().startswith("```"):
            if table:
                output.append("</tbody></table></div>")
                table = False
            output.append("</pre>" if fenced else "<pre>")
            fenced = not fenced
            continue
        if fenced:
            output.append(_escape(line) + "\n")
            continue
        if line.strip().startswith("|") and line.strip().endswith("|"):
            if re.fullmatch(r"[|\s:\-]+", line):
                continue
            if not table:
                output.append('<div class="table-scroll"><table><tbody>')
                table = True
            output.append(
                "<tr>"
                + "".join(
                    "<td>" + _escape(c.strip()) + "</td>"
                    for c in line.strip().strip("|").split("|")
                )
                + "</tr>"
            )
            continue
        if table:
            output.append("</tbody></table></div>")
            table = False
        if line.strip():
            heading = re.match(r"^#{1,6}\s+(.+)$", line)
            content = _escape(heading[1] if heading else line)
            content = re.sub(r"`([^`]+)`", r"<code>\1</code>", content)
            content = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", content)
            tag = "h3" if heading else "p"
            output.append(f"<{tag}>{content}</{tag}>")
    if fenced:
        output.append("</pre>")
    if table:
        output.append("</tbody></table></div>")
    return "\n".join(output)


def write_report(snapshot: dict, task_root: Path) -> dict:
    """Write static HTML in the isolated worker. No scripts or remote assets."""
    fr = snapshot["scope"]["language"] == "fr"

    def tr(french, english):
        return french if fr else english

    output = task_root / "investigation.html"
    project = task_root.parent.parent
    anchors = {r["node"]["id"]: f"cell-{i}" for i, r in enumerate(snapshot["records"])}
    with output.open("w", encoding="utf-8") as stream:

        def emit(value):
            stream.write(value)

        emit(
            '<!doctype html><html lang="'
            + ("fr" if fr else "en")
            + '"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1">'
            '<meta http-equiv="Content-Security-Policy" '
            "content=\"default-src 'none'; img-src data:; "
            "style-src 'unsafe-inline'\">"
            "<title>"
            + _escape(snapshot["project"]["name"])
            + "</title><style>"
            + _CSS
            + "</style></head><body><main>"
        )
        emit(
            '<header><p class="label">Linexcel · '
            + tr("Rapport d’investigation", "Investigation report")
            + "</p><h1>"
            + _escape(snapshot["project"]["name"])
            + "</h1><p>"
            + _escape(snapshot["project"]["filename"])
            + " · "
            + _when(snapshot["createdAt"])
            + "</p><p>"
            + tr(
                "Sélection explicite de cellules, avec un voisinage limité à "
                "200 nœuds et un niveau par cellule. Les preuves détaillées "
                "peuvent inclure des précédents supplémentaires issus des étapes "
                "de calcul enregistrées, avec leur propre couverture et pagination.",
                "Explicit cell selection, with a neighborhood limited to 200 "
                "nodes and one level per cell. Detailed evidence may include "
                "additional precedents from recorded calculation steps, with "
                "their own coverage and pagination.",
            )
            + '</p><nav aria-label="'
            + tr("Cellules sélectionnées", "Selected cells")
            + '"><ul>'
        )
        for node_id, anchor in anchors.items():
            emit('<li><a href="#' + anchor + '">' + _escape(node_id) + "</a></li>")
        emit(
            "</ul></nav><details><summary>"
            + tr("Identité des sources", "Source identity")
            + "</summary><p>SHA-256: <code>"
            + _escape(snapshot.get("sourceSha256"))
            + "</code></p><p>"
            + tr("Révision", "Revision")
            + ": <code>"
            + _escape(snapshot["project"]["revision"])
            + "</code></p></details></header>"
        )
        for record in snapshot["records"]:
            node = record["node"]
            calculation = record["calculation"]
            result = calculation["result"] if calculation else {}
            formula = node.get("sourceFormula")
            if not formula:
                formula = node.get("canonicalFormula") or node.get("formula")
                formula_label = tr(
                    "Formule dérivée / canonique", "Derived / canonical formula"
                )
            else:
                formula_label = tr("Formule source", "Source formula")
            emit(
                '<section id="'
                + anchors[node["id"]]
                + '"><h2>'
                + _escape(node["id"])
                + "</h2>"
            )
            if formula:
                emit(
                    "<h3>" + formula_label + "</h3><pre>" + _escape(formula) + "</pre>"
                )
            literal = node.get("kind") == "input"
            non_scalar = node.get("kind") == "range"
            emit(
                '<div class="values"><div class="value"><h3>'
                + tr("Valeur lue dans le fichier", "Value read from file")
                + '</h3><div class="number">'
                + (
                    tr("Plage de cellules", "Cell range")
                    if non_scalar
                    else _value(
                        node.get("value") if literal else node.get("cachedValue")
                    )
                )
                + '</div><p class="muted">'
                + (
                    tr(
                        "Aucune valeur scalaire pour cette plage",
                        "No scalar value for this range",
                    )
                    if non_scalar
                    else tr("Entrée source", "Source input")
                    if literal
                    else tr(
                        "Cache enregistré ; peut être absent ou ancien",
                        "Saved cache; may be absent or old",
                    )
                )
                + '</p></div><div class="value"><h3>'
                + tr("Calcul Linexcel", "Linexcel calculation")
                + "</h3>"
            )
            if calculation:
                status = result.get("status", "unknown")
                emit(
                    '<div class="number">'
                    + (
                        _value(result.get("value"))
                        if status == "completed"
                        else tr("Résultat indisponible", "Result unavailable")
                    )
                    + "</div><p>"
                    + _escape(status)
                    + '</p><p class="muted">'
                    + _when(calculation["createdAt"])
                    + "</p><p>"
                    + tr("Opération", "Operation")
                    + ": <code>"
                    + _escape(calculation["id"])
                    + "</code></p>"
                )
                for diagnostic in result.get("diagnostics", []):
                    emit("<p>" + _escape(diagnostic) + "</p>")
            else:
                emit(
                    "<p>"
                    + tr(
                        "Aucun calcul enregistré pour cette révision.",
                        "No calculation recorded for this revision.",
                    )
                    + "</p>"
                )
            attempt = record.get("calculationAttempt")
            if attempt and (not calculation or attempt["id"] != calculation["id"]):
                emit(
                    "<p>"
                    + tr("Dernière tentative : ", "Latest attempt: ")
                    + _escape(attempt["status"])
                    + " · "
                    + _when(attempt["createdAt"])
                    + "</p>"
                )
            emit(
                '</div></div><p class="muted">'
                + tr(
                    "Le cache, le résultat moteur et la documentation IA sont "
                    "des preuves distinctes. Un résultat moteur n’est pas une "
                    "certification Excel.",
                    "The cache, engine result and AI documentation are separate "
                    "evidence. An engine result is not Excel certification.",
                )
                + "</p>"
            )
            context = {n["id"]: n for n in record["context"]}
            for direction, label in (
                ("sources", tr("Sources directes", "Direct sources")),
                ("uses", tr("Usages directs", "Direct uses")),
            ):
                related = [
                    e["source"] if direction == "sources" else e["target"]
                    for e in record["edges"]
                    if e["target" if direction == "sources" else "source"] == node["id"]
                ]
                total = record["relationCounts"][direction]
                emit(
                    "<h3>"
                    + label
                    + "</h3><p>"
                    + str(len(related))
                    + " / "
                    + str(total)
                    + " "
                    + tr("relations incluses", "relations included")
                    + "</p>"
                )
                if related:
                    emit(
                        "<table><thead><tr><th>"
                        + tr("Cellule / plage", "Cell / range")
                        + "</th><th>"
                        + tr("Formule ou entrée source", "Formula or source input")
                        + "</th></tr></thead><tbody>"
                    )
                    for key in related:
                        other = context.get(key, node)
                        label_html = _escape(key)
                        if key in anchors:
                            label_html = (
                                '<a href="#' + anchors[key] + '">' + label_html + "</a>"
                            )
                        detail = other.get("sourceFormula") or other.get("formula")
                        if detail is None:
                            detail = (
                                other.get("value")
                                if other.get("kind") == "input"
                                else tr(
                                    "Plage / valeur non scalaire",
                                    "Range / non-scalar value",
                                )
                            )
                        emit(
                            "<tr><td>"
                            + label_html
                            + "</td><td>"
                            + _escape(detail)
                            + "</td></tr>"
                        )
                    emit("</tbody></table>")
            for range_record in record.get("ranges", []):
                emit(
                    "<details><summary>"
                    + _escape(range_record["id"])
                    + " · "
                    + str(len(range_record["members"]))
                    + " / "
                    + str(range_record["representedCellCount"])
                    + " "
                    + tr("cellules représentées incluses", "represented cells included")
                    + "</summary><p>"
                    + str(range_record["positionCount"])
                    + " "
                    + tr(
                        "positions dans la plage ; les cellules absentes "
                        "de l’index ne sont pas matérialisées.",
                        "range positions; cells absent from the index "
                        "are not materialized.",
                    )
                    + "</p><table><thead><tr><th>"
                    + tr("Cellule", "Cell")
                    + "</th><th>"
                    + tr("Entrée / formule source", "Source input / formula")
                    + "</th></tr></thead><tbody>"
                )
                for member in range_record["members"]:
                    detail = member.get("sourceFormula") or member.get("formula")
                    if detail is None:
                        detail = member.get("value")
                    emit(
                        "<tr><td>"
                        + _escape(member["id"])
                        + "</td><td>"
                        + _escape(detail)
                        + "</td></tr>"
                    )
                emit("</tbody></table></details>")
            doc = record["documentation"]
            emit(
                "<h3>"
                + tr("Documentation IA enregistrée", "Saved AI documentation")
                + "</h3>"
            )
            if doc:
                content = doc["result"]
                emit(
                    '<p class="muted">'
                    + _escape(content.get("model"))
                    + " · "
                    + _when(doc["createdAt"])
                    + "</p>"
                )
                if (
                    calculation
                    and content.get("sources", {}).get("calculationTaskId")
                    != calculation["id"]
                ):
                    emit(
                        "<p>"
                        + tr(
                            "Cette documentation n’est pas liée au calcul affiché.",
                            "This documentation is not bound to "
                            "the displayed calculation.",
                        )
                        + "</p>"
                    )
                emit(
                    '<div class="prose">'
                    + _prose(content.get("markdown", ""))
                    + "</div>"
                )
                for warning in content.get("warnings", []):
                    emit('<p class="muted">' + _escape(warning) + "</p>")
            else:
                emit(
                    "<p>"
                    + tr(
                        "Aucune documentation enregistrée dans cette langue "
                        "pour cette révision.",
                        "No documentation saved in this language for this revision.",
                    )
                    + "</p>"
                )
            attempt = record.get("documentationAttempt")
            if attempt and (not doc or attempt["id"] != doc["id"]):
                emit(
                    "<p>"
                    + tr(
                        "Dernière tentative de documentation : ",
                        "Latest documentation attempt: ",
                    )
                    + _escape(attempt["status"])
                    + " · "
                    + _when(attempt["createdAt"])
                    + "</p>"
                )
            proof = {
                "source": node,
                "calculation": calculation,
                "calculationAttempt": record.get("calculationAttempt"),
                "documentationAttempt": record.get("documentationAttempt"),
                "bounds": record["bounds"],
            }
            emit(
                "<details><summary>"
                + tr(
                    "Preuves et couverture détaillées", "Detailed evidence and coverage"
                )
                + "</summary><pre>"
                + _escape(json.dumps(proof, ensure_ascii=False, indent=2))
                + "</pre></details></section>"
            )
        for shot in snapshot["captures"]:
            path = (project / shot["path"]).resolve()
            path.relative_to((project / "captures").resolve())
            if file_sha256(path) != shot["sha256"]:
                raise ValueError("Capture changed after report snapshot")
            emit(
                "<section><h2>"
                + _escape(shot["sheet"])
                + "</h2><p>"
                + tr(
                    "Rendu LibreOffice ; ses valeurs peuvent différer du cache source.",
                    "LibreOffice rendering; its values may differ "
                    "from the source cache.",
                )
                + '</p><img alt="'
                + _escape(shot["name"])
                + '" src="data:image/png;base64,'
            )
            with path.open("rb") as source:
                while chunk := source.read(3 * 16384):
                    emit(base64.b64encode(chunk).decode("ascii"))
            emit('"><p>SHA-256: <code>' + shot["sha256"] + "</code></p></section>")
        emit("</main></body></html>")
    return {
        "format": "html",
        "file": output.name,
        "reportVersion": 1,
        "sha256": file_sha256(output),
        "bytes": output.stat().st_size,
        "revision": snapshot["project"]["revision"],
        "scope": snapshot["scope"],
    }
