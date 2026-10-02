"""Private web task entry point; the parent establishes resource isolation."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path


def main():
    root = Path(sys.argv[1])
    request = json.loads((root / "request.json").read_text("utf-8"))
    project = root.parent.parent
    try:
        if request["operation"] == "import":
            from linexcel.lazy_store import build_index

            result = build_index(
                project / "workbook.xlsx",
                root / "index.sqlite",
                project / "references",
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

            shots = render_workbook_screenshots(
                project / "workbook.xlsx",
                "workbook.xlsx",
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
        else:
            output = {"error": {"kind": type(exc).__name__, "message": str(exc)}}
    path = root / "result.json"
    path.write_text(json.dumps(output, ensure_ascii=False, allow_nan=False), "utf-8")


if __name__ == "__main__":
    main()
