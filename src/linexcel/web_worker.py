"""Private web task entry point; the parent establishes resource isolation."""

from __future__ import annotations

import base64
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
            from linexcel.lazy import build_structure

            result = build_structure(project / "workbook.xlsx", project / "references")
        elif request["operation"] == "evaluate":
            from linexcel.lazy import evaluate_node

            result = evaluate_node(
                project / "workbook.xlsx", request["nodeId"], project / "references"
            )
        elif request["operation"] == "capture":
            from linexcel.insights import render_workbook_screenshots

            shots = render_workbook_screenshots(
                (project / "workbook.xlsx").read_bytes(),
                "workbook.xlsx",
                root / "captures",
            )
            mapping = shots if isinstance(shots, dict) else {"Pages imprimées": shots}
            result = {
                "screenshots": [
                    {
                        "sheet": sheet,
                        "name": path.name,
                        "data": "data:image/png;base64,"
                        + base64.b64encode(path.read_bytes()).decode(),
                    }
                    for sheet, paths in mapping.items()
                    for path in paths
                ],
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
