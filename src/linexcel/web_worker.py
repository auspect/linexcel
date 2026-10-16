"""Private web task entry point; the parent establishes resource isolation."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any


def main():
    root = Path(sys.argv[1])
    request = json.loads((root / "request.json").read_text("utf-8"))
    project = root.parent.parent
    metadata = json.loads((project / "project.json").read_text("utf-8"))
    output: dict[str, Any]
    try:
        if request["operation"] == "import":
            from linexcel.lazy_store import build_index

            result = build_index(
                project / "workbook.xlsx",
                root / "index.sqlite",
                project / "references",
                filename=metadata["filename"],
            )
        elif request["operation"] == "build_graph":
            from linexcel.graph_store import build_graph

            result = build_graph(project / "index.sqlite", root / "hierarchy.sqlite")
        elif request["operation"] == "evaluate":
            from linexcel.lazy import evaluate_node

            result = evaluate_node(
                project / "workbook.xlsx",
                request["nodeId"],
                project / "references",
                index_path=project / "index.sqlite",
            )
        elif request["operation"] == "capture":
            from linexcel.insights import render_workbook_screenshots

            if metadata["filename"].lower().endswith((".xla", ".xlam")):
                raise ValueError(
                    "An add-in is a static macro report; "
                    "worksheet capture is unavailable."
                )
            shots = render_workbook_screenshots(
                project / "workbook.xlsx",
                metadata["filename"],
                root / "captures",
            )
            mapping = shots if isinstance(shots, dict) else {"Pages imprimées": shots}
            published = root / "capture-files"
            published.mkdir()
            screenshots = []
            for sheet, paths in mapping.items():
                for path in paths:
                    name = path.name
                    filename = f"{len(screenshots)}.png"
                    path.replace(published / filename)
                    shot = {"sheet": sheet, "name": name, "file": filename}
                    try:
                        from PIL import Image

                        thumbnail = f"{len(screenshots)}.thumb.png"
                        with Image.open(published / filename) as image:
                            image.thumbnail((320, 240))
                            image.save(published / thumbnail)
                        shot["thumbnail"] = thumbnail
                    except Exception:
                        # The full image remains usable if thumbnail creation fails.
                        pass
                    screenshots.append(shot)
            result = {
                "screenshots": screenshots,
                "notice": "Rendu LibreOffice, qui peut recalculer "
                "les valeurs affichées. "
                "Le fichier source et son cache restent inchangés. Aucune IA.",
            }
        elif request["operation"] in {"document", "describe_capture"}:
            from linexcel.web_ai import WebAIError, generate

            secret = os.environ.pop("LINEXCEL_WEB_AI_CONFIG", None)
            if not secret:
                raise WebAIError("ai_configuration", "Aucun fournisseur IA configuré.")
            snapshot = json.loads((root / "ai_input.json").read_text("utf-8"))
            result = generate(
                snapshot, json.loads(secret), request["options"], request["tokens"]
            )
        else:
            raise ValueError("Unknown operation")
        output = {"result": result}
    except Exception as exc:
        if request["operation"] in {"document", "describe_capture"}:
            from linexcel.web_ai import WebAIError

            review_evidence = None
            if isinstance(exc, WebAIError) and exc.review_report is not None:
                import hashlib

                # Keep prompts and returned draft text in the owner-scoped task
                # directory, not in coordinator metadata or a public error body.
                proof = json.dumps(exc.review_report, ensure_ascii=False).encode()
                (root / "ai-review-failed.json").write_bytes(proof)
                review_evidence = {
                    "file": "ai-review-failed.json",
                    "sha256": hashlib.sha256(proof).hexdigest(),
                    "scope": "private owner-scoped task evidence",
                }
            output = (
                {"error": {"kind": exc.kind, "message": str(exc), "usage": exc.usage}}
                if isinstance(exc, WebAIError)
                else {
                    "error": {
                        "kind": "ai_failure",
                        "message": "Échec de l’opération IA. "
                        "Vérifiez la configuration serveur.",
                    }
                }
            )
            if review_evidence is not None:
                output["error"]["reviewEvidence"] = review_evidence
        else:
            output = {"error": {"kind": type(exc).__name__, "message": str(exc)}}
            capture_error: dict[str, Any] = output["error"]
            if request["operation"] == "capture":
                from linexcel.insights import WorkbookRenderError

                if isinstance(exc, WorkbookRenderError) and exc.exit_code is not None:
                    from linexcel.execution import _failure_details

                    failure = _failure_details("crashed", exc.exit_code, "")
                    capture_error.update(renderer=exc.renderer, exitCode=exc.exit_code)
                    if failure["kind"] != "unknown_exit":
                        capture_error.update(
                            kind=failure["kind"],
                            message=f"{exc}. {failure['summary']}",
                        )
                    if failure["kind"] == "memory_limit":
                        memory = request.get("memoryMb")
                        budget = f" de {memory} Mio" if memory is not None else ""
                        capture_error["message"] = (
                            f"Budget mémoire{budget} insuffisant pour le rendu "
                            f"des captures. Augmentez le budget puis réessayez. {exc}"
                        )
    path = root / "result.json"
    path.write_text(json.dumps(output, ensure_ascii=False, allow_nan=False), "utf-8")


if __name__ == "__main__":
    main()
