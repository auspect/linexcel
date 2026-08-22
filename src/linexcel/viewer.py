"""Standalone HTML viewer for the lineage graph.

Generates a complete, **offline** HTML document: Cytoscape libraries are
embedded in the file (``assets/`` folder), so no network access is required.
It is meant to be:

- saved and opened in a browser (``result.save_html(...)``);
- displayed inline in marimo / Jupyter via ``result._repr_html_()`` which
  wraps it in an isolated ``<iframe srcdoc>``.

No dependency on FastAPI or the Svelte frontend: rendering lives entirely in
this file. If assets are missing, a CDN fallback is used (requires network).
"""

from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any

from linexcel.i18n import LANGUAGES, ui_payload, validate_language

_ASSETS = Path(__file__).parent / "assets"

#: UI languages the viewer can render. Alias of :data:`linexcel.i18n.LANGUAGES`,
#: which is the single source of truth.
SUPPORTED_LANGUAGES = LANGUAGES

# Template placeholders, substituted in a single pass so that a value can never
# be rescanned: chained str.replace calls let the graph JSON — i.e. arbitrary
# workbook content — be rewritten by a later placeholder.
_PLACEHOLDER_RE = re.compile(
    "__GRAPH_JSON__|__I18N_JSON__|__TITLE__|__LANG__|__SHEET_OPTIONS__"
)

# Load order (UMD chain): cytoscape, then the layout stack.
_ASSET_FILES = (
    "cytoscape.min.js",
    "layout-base.js",
    "cose-base.js",
    "cytoscape-fcose.js",
    "cytoscape-dagre.min.js",
)

# CDN fallback if assets are not embedded (network access needed).
_CDN = (
    "https://cdn.jsdelivr.net/npm/cytoscape@3.30.2/dist/cytoscape.min.js",
    "https://cdn.jsdelivr.net/npm/layout-base@2.0.1/layout-base.js",
    "https://cdn.jsdelivr.net/npm/cose-base@2.2.0/cose-base.js",
    "https://cdn.jsdelivr.net/npm/cytoscape-fcose@2.2.0/cytoscape-fcose.js",
    "https://cdn.jsdelivr.net/npm/cytoscape-dagre@3.0.0/cytoscape-dagre.min.js",
)


@lru_cache(maxsize=1)
def _inline_scripts() -> str | None:
    """Concatenate embedded JS into inline <script> tags (or ``None``)."""
    if not all((_ASSETS / f).exists() for f in _ASSET_FILES):
        return None
    parts = []
    for fname in _ASSET_FILES:
        code = (_ASSETS / fname).read_text(encoding="utf-8")
        code = code.replace("</script>", "<\\/script>")  # prevent tag closure
        parts.append(f"<script>{code}</script>")
    return "\n".join(parts)


def _safe_json(obj: Any) -> str:
    """JSON embeddable in a <script> tag.

    Escapes ``<`` (to prevent closing the tag or opening a ``<!--``) and
    U+2028/U+2029, line terminators in JavaScript that would break the
    literal.
    """
    blob = json.dumps(obj, ensure_ascii=False, default=str)
    return (
        blob.replace("<", "\\u003c")
        .replace(chr(0x2028), "\\u2028")
        .replace(chr(0x2029), "\\u2029")
    )


def wrap_iframe(document_html: str, height: int = 640) -> str:
    """Wrap an HTML document in an isolated iframe (for notebooks).

    The document (including embedded Cytoscape) is base64-encoded in a
    ``data:`` URI: no costly attribute escaping, and the iframe is isolated
    in an opaque origin (``allow-scripts`` allows the embedded JS).
    """
    import base64

    b64 = base64.b64encode(document_html.encode("utf-8")).decode("ascii")
    return (
        f'<iframe src="data:text/html;base64,{b64}" '
        f'style="width:100%;height:{height}px;border:1px solid #e1e0d9;'
        f'border-radius:8px;background:#fcfcfb" '
        f'sandbox="allow-scripts"></iframe>'
    )


def render_html(
    graph: dict[str, Any],
    title: str = "Lineage Excel",
    full_document: bool = True,
    language: str = "en",
) -> str:
    """Build the viewer HTML for a given graph.

    ``language`` must be one of :data:`SUPPORTED_LANGUAGES`; it is interpolated
    into a JavaScript string literal and an HTML attribute, so it is validated
    against that closed set rather than escaped.
    """
    validate_language(language)
    substitutions = {
        "__GRAPH_JSON__": _safe_json(graph),
        "__I18N_JSON__": _safe_json(ui_payload(language)),
        "__TITLE__": _escape_text(title),
        "__LANG__": language,
        "__SHEET_OPTIONS__": _sheet_options(graph, language),
    }
    # A callable replacement keeps the substituted text literal: no backslash
    # or \g<n> expansion, and no placeholder inside workbook data is rescanned.
    body = _PLACEHOLDER_RE.sub(lambda m: substitutions[m.group(0)], _TEMPLATE)
    if not full_document:
        return body
    scripts = _inline_scripts()
    if scripts is None:
        scripts = "\n".join(f'<script src="{url}"></script>' for url in _CDN)
    return (
        f"<!doctype html><html lang='{language}'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width, initial-scale=1'>"
        # Light is the shipped default whatever the OS prefers — spreadsheet
        # work is done in light. The theme toggle rewrites this content
        # attribute so the UA keeps painting its own surfaces (scrollbars,
        # overscroll) to match whichever theme is actually on screen.
        "<meta name='color-scheme' content='light'>"
        f"<title>{_escape_text(title)}</title>{scripts}</head>"
        f"<body>{body}</body></html>"
    )


def _sheet_options(graph: dict[str, Any], language: str) -> str:
    """``<option>`` markup for the sheet filter.

    The sheet list is derived from the nodes themselves — the graph carries no
    sheet index — keeping first-seen order so the dropdown follows the workbook
    rather than an alphabet. Nodes without a sheet (opaque references, names,
    VBA) contribute no entry: they are reachable through the cross-sheet
    neighbourhood, not as a filter target of their own.
    """
    seen: list[str] = []
    for node in graph.get("nodes", []):
        sheet = node.get("sheet")
        if isinstance(sheet, str) and sheet and sheet not in seen:
            seen.append(sheet)
    strings = ui_payload(language)
    all_label = strings.get(language, strings["en"])["all_sheets"]
    options = [f'<option value="__all__">{_escape_text(all_label)}</option>']
    options += [
        f'<option value="{_escape_text(s)}">{_escape_text(s)}</option>' for s in seen
    ]
    return "".join(options)


def _escape_text(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


# Body: styles + containers + JS logic. Palette from the validated CVD-safe
# dataviz design system. Colors follow node TYPE, never recycled.
#
# It lives in assets/viewer.html rather than in a string here: 1,700 lines of
# HTML, CSS and JavaScript inside Python are invisible to every tool that
# could check them — no highlighting, no linting, no formatter — and every
# edit means reading markup through Python's quoting rules. The file is read
# verbatim, exactly as the raw string was.
_TEMPLATE = (_ASSETS / "viewer.html").read_text(encoding="utf-8")
