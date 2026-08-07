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
_TEMPLATE = r"""
<style>
  .lin-root {
    --surface: #fcfcfb; --ink: #0b0b0b; --ink2: #52514e; --muted: #898781;
    --line: #c3c2b7; --hair: #e1e0d9; --blue: #2a78d6;
    font-family: system-ui, -apple-system, 'Segoe UI', sans-serif;
    color: var(--ink); position: absolute; inset: 0; display: flex;
    flex-direction: column; background: var(--surface); overflow: hidden;
  }
  .lin-root * { box-sizing: border-box; }
  .lin-bar {
    display: flex; align-items: center; gap: .5rem; padding: .45rem .7rem;
    border-bottom: 1px solid var(--hair); background: var(--surface);
    flex-wrap: wrap;
  }
  .lin-bar h1 { font-size: .95rem; margin: 0 .5rem 0 0; }
  .lin-bar .stat { color: var(--ink2); font-size: .8rem; }
  .lin-bar input {
    padding: .3rem .5rem; border: 1px solid var(--line); border-radius: 6px;
    font: inherit; width: 200px;
  }
  .lin-bar button {
    padding: .3rem .6rem; border: 1px solid var(--line); border-radius: 6px;
    background: #fff; cursor: pointer; font: inherit;
  }
  .lin-bar select {
    padding: .3rem .5rem; border: 1px solid var(--line); border-radius: 6px;
    background: #fff; cursor: pointer; font: inherit; max-width: 200px;
  }
  .lin-bar select:disabled { cursor: default; opacity: .5; }
  .lin-bar button.active { background: var(--ink); color: #fff; }
  .lin-tab[hidden] { display: none; }
  .lin-main { flex: 1; display: flex; min-height: 0; }
  .lin-main.hidden { display: none; }
  .lin-cy { flex: 1; position: relative; min-width: 0; }
  .lin-legend {
    position: absolute; left: .6rem; bottom: .6rem; display: flex; gap: .6rem;
    flex-wrap: wrap; background: rgba(252,252,251,.92);
    border: 1px solid var(--hair); border-radius: 8px; padding: .35rem .6rem;
    font-size: .72rem; color: var(--ink2); max-width: 72%;
  }
  .lin-legend span { display: inline-flex; align-items: center; gap: .3rem; }
  .lin-sw { width: 11px; height: 11px; border-radius: 3px; display: inline-block; }
  .lin-panel {
    width: 340px; flex-shrink: 0; overflow-y: auto; padding: .8rem .9rem 2rem;
    border-left: 1px solid var(--hair); background: var(--surface);
    font-size: .85rem;
  }
  .lin-panel.hidden { display: none; }
  .lin-panel h2 { font-size: .95rem; margin: .3rem 0 .5rem; word-break: break-all; }
  .lin-panel h3 {
    font-size: .72rem; text-transform: uppercase; letter-spacing: .04em;
    color: var(--muted); margin: .8rem 0 .3rem;
  }
  .lin-badge { color: #fff; font-size: .7rem; padding: .12rem .5rem; border-radius: 99px; }
  .lin-formula {
    display: block; background: #f0efec; border-radius: 6px; padding: .45rem .55rem;
    font-family: ui-monospace, monospace; font-size: .8rem; white-space: pre-wrap;
    word-break: break-all;
  }
  .lin-val { font-size: 1.1rem; font-weight: 600; }
  .lin-src {
    display: inline-block; vertical-align: middle; margin-left: .4rem;
    font-size: .62rem; font-weight: 700; color: #fff; background: var(--blue);
    border-radius: 99px; padding: .08rem .45rem; letter-spacing: .01em;
  }
  .lin-cmp { font-size: .74rem; color: var(--muted); margin-top: .15rem; }
  .lin-step {
    border-left: 3px solid var(--blue); background: #fff; border-radius: 0 6px 6px 0;
    padding: .3rem .5rem; margin: .35rem 0; box-shadow: 0 1px 2px rgba(11,11,11,.05);
  }
  .lin-step.lin-final { border-left-color: #1baf7a; }
  .lin-step.lin-final .lin-op { background: #1baf7a; }
  .lin-step.lin-final .lin-val { margin-bottom: .2rem; }
  .lin-op {
    font-size: .68rem; font-weight: 700; color: #fff; background: var(--blue);
    border-radius: 4px; padding: .04rem .35rem; margin-right: .35rem;
  }
  .lin-expr { font-family: ui-monospace, monospace; font-size: .74rem; }
  .lin-in {
    display: inline-block; font-size: .7rem; background: #e8f6f0;
    border: 1px solid #bfe7d8; border-radius: 4px; padding: .03rem .3rem;
    margin: .15rem .15rem 0 0; font-family: ui-monospace, monospace;
  }
  .lin-in.lit { background: #f0efec; border-color: var(--hair); }
  .lin-nav { list-style: none; margin: 0; padding: 0; }
  .lin-nav li { display: flex; justify-content: space-between; gap: .4rem; align-items: center; }
  .lin-nav button {
    border: none; background: none; color: var(--blue); cursor: pointer;
    text-align: left; padding: .1rem 0; display: inline-flex; align-items: center;
    gap: .3rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font: inherit;
  }
  .lin-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
  .lin-ek { font-size: .66rem; color: var(--muted); flex-shrink: 0; }
  .lin-code { background: #f0efec; border-radius: 6px; padding: .5rem; font-size: .72rem;
    overflow-x: auto; max-height: 260px; font-family: ui-monospace, monospace; }
  .lin-doc { font-size: .82rem; line-height: 1.45; }
  .lin-doc h1,.lin-doc h2,.lin-doc h3 { font-size: .9rem; margin: .4rem 0 .2rem; }
  .lin-doc p { margin: .25rem 0; }
  .lin-doc ul,.lin-doc ol { padding-left: 1.2rem; margin: .25rem 0; }
  .lin-doc code { background: #f0efec; padding: .1rem .3rem; border-radius: 3px;
    font-family: ui-monospace, monospace; font-size: .78rem; }
  .lin-doc strong { font-weight: 600; }
  .lin-hint { color: var(--muted); font-size: .76rem; margin: .2rem 0; }
  .lin-close { margin-left: auto; border: none; background: none; cursor: pointer; font-size: 1rem; }
  .lin-fallback { position: absolute; inset: 0; display: flex; align-items: center;
    justify-content: center; text-align: center; padding: 2rem; color: var(--ink2); }
  .lin-overview { flex: 1; overflow-y: auto; padding: 1.4rem clamp(1rem, 4vw, 4rem) 3rem; }
  .lin-overview-inner { max-width: 880px; }
  .lin-overview h2 { font-size: 1.1rem; margin: 0 0 .8rem; }
  .lin-overview .lin-doc { font-size: .92rem; line-height: 1.55; }
  .lin-ai-box {
    background: #fbf9ff; border: 1px solid #e5dbff; border-radius: 8px;
    padding: .75rem; margin: .8rem 0;
  }
  .lin-ai-badge {
    display: inline-block; font-size: .65rem; font-weight: 700;
    color: #5f3dc4; background: #efe9ff; border: 1px solid #d0bfff;
    border-radius: 4px; padding: .1rem .35rem; margin-bottom: .4rem;
    text-transform: uppercase; letter-spacing: .03em;
  }
  #lin-sheets-sidebar button {
    display: block;
    width: 100%;
    text-align: left;
    background: none;
    border: 1px solid transparent;
    border-radius: 6px;
    padding: .5rem .8rem;
    font-size: .85rem;
    font-weight: 500;
    color: var(--ink2);
    cursor: pointer;
    transition: all .15s ease;
  }
  #lin-sheets-sidebar button:hover {
    background: var(--hair);
    color: var(--ink);
  }
  #lin-sheets-sidebar button.active {
    background: var(--ink) !important;
    color: #fff !important;
    font-weight: 600;
  }
</style>

<div class="lin-root">
  <div class="lin-bar">
    <h1>__TITLE__</h1>
    <span class="stat" id="lin-stats"></span>
    <button id="lin-tab-graph" class="lin-tab active">Graph</button>
    <button id="lin-tab-overview" class="lin-tab" hidden>Workbook overview</button>
    <button id="lin-tab-sheets" class="lin-tab" hidden>Sheets</button>
    <button id="lin-tab-screenshots" class="lin-tab" hidden>Visual preview</button>
    <span style="flex:1"></span>
    <input id="lin-search" placeholder="Search… ⏎" />
    <select id="lin-sheet-filter" title="Sheet filter">__SHEET_OPTIONS__</select>
    <button id="lin-lay-dagre">Flow</button>
    <button id="lin-lay-fcose" class="active">Organic</button>
    <button id="lin-zoom-in" title="Zoom In">+</button>
    <button id="lin-zoom-out" title="Zoom Out">−</button>
    <button id="lin-fit" title="Fit All">Fit</button>
    <button id="lin-fit-sel" style="display: none;" title="Fit Selection">Fit Selection</button>
  </div>
  <div class="lin-main" id="lin-graph-main">
    <div class="lin-cy" id="lin-cy">
      <div class="lin-legend" id="lin-legend"></div>
    </div>
    <aside class="lin-panel hidden" id="lin-panel"></aside>
  </div>
  <div class="lin-main hidden" id="lin-overview-main">
    <article class="lin-overview"><div class="lin-overview-inner" id="lin-overview"></div></article>
  </div>
  <div class="lin-main hidden" id="lin-sheets-main">
    <div id="lin-sheets-container" style="display: flex; flex: 1; overflow: hidden; height: 100%;">
      <aside id="lin-sheets-sidebar"></aside>
      <article style="flex: 1; overflow-y: auto; padding: 1.5rem clamp(1rem, 4vw, 4rem) 3rem; background: #fff;">
        <div id="lin-sheet-details" class="lin-overview-inner"></div>
      </article>
    </div>
  </div>
  <div class="lin-main hidden" id="lin-screenshots-main">
    <article class="lin-overview"><div class="lin-overview-inner" id="lin-screenshots"></div></article>
  </div>
</div>

<script>
(function () {
  var GRAPH = __GRAPH_JSON__;
  var LANG = "__LANG__";
  var I18N = __I18N_JSON__;

  function _t(key, replacements) {
    var langDict = I18N[LANG] || I18N.en;
    var str = langDict[key] || I18N.en[key] || key;
    if (replacements) {
      Object.keys(replacements).forEach(function (k) {
        str = str.replace('{' + k + '}', replacements[k]);
      });
    }
    return str;
  }

  var KIND = {
    cell:  { color: '#2a78d6', shape: 'round-rectangle', labelKey: 'kind_cell' },
    group: { color: '#4a3aa7', shape: 'round-rectangle', labelKey: 'kind_group' },
    input: { color: '#1baf7a', shape: 'ellipse', labelKey: 'kind_input' },
    name:  { color: '#eda100', shape: 'diamond', labelKey: 'kind_name' },
    vba:   { color: '#eb6834', shape: 'hexagon', labelKey: 'kind_vba' },
    misc:  { color: '#898781', shape: 'octagon', labelKey: 'kind_misc' },
    opaque:{ color: '#898781', shape: 'ellipse', labelKey: 'kind_opaque' }
  };
  // Value provenance: where a displayed value comes from. Colors reuse the
  // node palette — blue = linexcel engine, green = the file, orange = fallback.
  var SRC = {
    engine:   { labelKey: 'value_recalc',    color: 'var(--blue)' },
    file:     { labelKey: 'value_from_file', color: '#1baf7a' },
    fallback: { labelKey: 'value_fallback',  color: '#eda100' }
  };
  // Engine error objects → the text Excel itself would show.
  var ERRTEXT = {
    Div: '#DIV/0!', Ref: '#REF!', Value: '#VALUE!', NA: '#N/A',
    Name: '#NAME?', Num: '#NUM!', Null: '#NULL!'
  };
  var byId = {};
  GRAPH.nodes.forEach(function (n) { byId[n.id] = n; });
  // Sheet filter state: ALL_SHEETS mirrors the sentinel option value rendered
  // server-side, CUR_SHEET is what the graph is currently narrowed to.
  var ALL_SHEETS = '__all__';
  var CUR_SHEET = ALL_SHEETS;

  function boot() {
    setupTabs();

    // Set UI labels in the correct language
    document.getElementById('lin-tab-graph').textContent = _t('graph');
    document.getElementById('lin-tab-overview').textContent = _t('overview');
    document.getElementById('lin-tab-sheets').textContent = _t('sheets_tab');
    document.getElementById('lin-tab-screenshots').textContent = _t('visual');
    document.getElementById('lin-search').placeholder = _t('search');
    document.getElementById('lin-zoom-in').title = _t('zoom_in');
    document.getElementById('lin-zoom-out').title = _t('zoom_out');
    document.getElementById('lin-fit').title = _t('fit_all');
    document.getElementById('lin-fit').textContent = _t('fit_all');
    document.getElementById('lin-fit-sel').title = _t('fit_sel');
    document.getElementById('lin-fit-sel').textContent = _t('fit_sel');
    var sheetSel = document.getElementById('lin-sheet-filter');
    sheetSel.title = _t('sheet_filter');
    sheetSel.setAttribute('aria-label', _t('sheet_filter'));

    var cyContainer = document.getElementById('lin-cy');
    if (typeof cytoscape === 'undefined') {
      var f = document.createElement('div');
      f.className = 'lin-fallback';
      f.textContent = _t('fallback');
      cyContainer.appendChild(f);
      sheetSel.disabled = true;  // nothing to filter without a rendered graph
      return;
    }
    // Register layout extensions if present.
    var hasFcose = false;
    try {
      if (window.cytoscapeFcose) { cytoscape.use(window.cytoscapeFcose); hasFcose = true; }
      if (window.cytoscapeDagre) { cytoscape.use(window.cytoscapeDagre); }
    } catch (e) { /* already registered */ }

    var stats = GRAPH.meta.stats;
    var vbaText = stats.vbaProcs ? ' · ' + stats.vbaProcs + ' VBA' : '';
    document.getElementById('lin-stats').textContent = _t('stats', {
      formulas: stats.totalFormulas.toLocaleString(LANG),
      nodes: stats.totalNodes,
      edges: stats.totalEdges,
      vba: vbaText
    });

    var big = GRAPH.nodes.length + GRAPH.edges.length > 2500;
    var degreeMap = {};
    GRAPH.edges.forEach(function (e) {
      degreeMap[e.source] = (degreeMap[e.source] || 0) + 1;
      degreeMap[e.target] = (degreeMap[e.target] || 0) + 1;
    });
    var elements = [];
    GRAPH.nodes.forEach(function (n) {
      elements.push({ data: {
        id: n.id, label: shortLabel(n), size: nodeSize(n),
        degree: degreeMap[n.id] || 0, sheet: n.sheet || ''
      }, classes: n.kind });
    });
    GRAPH.edges.forEach(function (e) {
      elements.push({ data: { id: e.id, source: e.source, target: e.target },
        classes: e.kind + (e.approx ? ' approx' : '') });
    });

    var hasDagre = typeof dagre !== 'undefined' || window.cytoscapeDagre;
    var initial = hasFcose ? 'fcose' : (hasDagre ? 'dagre' : 'cose');
    var curLayout = initial;
    var cy = cytoscape({
      container: cyContainer, elements: elements,
      minZoom: 0.05, maxZoom: 4, wheelSensitivity: 0.75,
      pixelRatio: big ? 1 : 'auto', textureOnViewport: big, hideEdgesOnViewport: big,
      style: buildStyle(big), layout: layoutOpts(initial, hasFcose)
    });

    cy.on('tap', 'node', function (ev) { select(cy, ev.target.id()); });
    cy.on('tap', function (ev) { if (ev.target === cy) clearSel(cy); });
    renderPlaceholder();

    document.getElementById('lin-fit').onclick = function () {
      cy.animate({ fit: { padding: 40 }, duration: 250 });
    };
    document.getElementById('lin-zoom-in').onclick = function () {
      var currentZoom = cy.zoom();
      var newZoom = Math.min(currentZoom * 1.4, cy.maxZoom());
      cy.zoom({
        level: newZoom,
        renderedPosition: { x: cyContainer.clientWidth / 2, y: cyContainer.clientHeight / 2 }
      });
    };
    document.getElementById('lin-zoom-out').onclick = function () {
      var currentZoom = cy.zoom();
      var newZoom = Math.max(currentZoom / 1.4, cy.minZoom());
      cy.zoom({
        level: newZoom,
        renderedPosition: { x: cyContainer.clientWidth / 2, y: cyContainer.clientHeight / 2 }
      });
    };
    var fitSelBtn = document.getElementById('lin-fit-sel');
    fitSelBtn.onclick = function () {
      var sel = cy.nodes(':selected');
      if (sel.length > 0) {
        cy.animate({ fit: { eles: sel, padding: 40 }, duration: 250 });
      }
    };
    cy.on('select unselect', function () {
      var sel = cy.nodes(':selected');
      if (sel.length > 0) {
        fitSelBtn.style.display = 'inline-block';
      } else {
        fitSelBtn.style.display = 'none';
      }
    });

    var bd = document.getElementById('lin-lay-dagre');
    var bf = document.getElementById('lin-lay-fcose');
    bd.onclick = function () {
      setActive(bd, bf); curLayout = 'dagre';
      cy.elements(':visible').layout(layoutOpts('dagre', hasFcose)).run();
    };
    bf.onclick = function () {
      setActive(bf, bd); curLayout = 'fcose';
      cy.elements(':visible').layout(layoutOpts('fcose', hasFcose)).run();
    };
    if (initial === 'dagre') setActive(bd, bf);
    sheetSel.onchange = function () {
      applySheetFilter(cy, sheetSel.value, curLayout, hasFcose);
    };
    var search = document.getElementById('lin-search');
    search.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter') return;
      var q = search.value.trim().toLowerCase();
      if (!q) return;
      var matched = GRAPH.nodes.filter(function (n) {
        return (n.label || '').toLowerCase().indexOf(q) >= 0
          || (n.formula || '').toLowerCase().indexOf(q) >= 0;
      });
      search.title = _t('search_matches', { count: matched.length });
      if (!matched.length) return;
      // The first hit drives the panel and the neighbourhood highlight; every
      // other match joins the selection and is undimmed, then the view frames
      // the whole set.
      select(cy, matched[0].id);
      var eles = cy.collection();
      cy.batch(function () {
        for (var i = 0; i < matched.length; i++) {
          var ele = cy.getElementById(matched[i].id);
          ele.select(); ele.removeClass('dimmed');
          eles = eles.union(ele);
        }
      });
      cy.animate({ fit: { eles: eles, padding: 40 }, duration: 300 });
    });
    buildLegend();
    setupScreenshots();
  }

  function switchTab(activeBtn, activeMain) {
    var graphButton = document.getElementById('lin-tab-graph');
    var overviewButton = document.getElementById('lin-tab-overview');
    var sheetsButton = document.getElementById('lin-tab-sheets');
    var screenshotsButton = document.getElementById('lin-tab-screenshots');

    var graphMain = document.getElementById('lin-graph-main');
    var overviewMain = document.getElementById('lin-overview-main');
    var sheetsMain = document.getElementById('lin-sheets-main');
    var screenshotsMain = document.getElementById('lin-screenshots-main');

    var allBtns = [graphButton, overviewButton, sheetsButton, screenshotsButton];
    var allMains = [graphMain, overviewMain, sheetsMain, screenshotsMain];

    allBtns.forEach(function(btn) { btn.classList.remove('active'); });
    allMains.forEach(function(m) { m.classList.add('hidden'); });

    activeBtn.classList.add('active');
    activeMain.classList.remove('hidden');
  }

  function setupScreenshots() {
    var images = GRAPH.meta.screenshots;
    if (!images || !images.length) return;
    if (typeof images === 'object' && !Array.isArray(images)) return;

    var btn = document.getElementById('lin-tab-screenshots');
    btn.hidden = false;

    var container = document.getElementById('lin-screenshots');
    container.appendChild(el('h2', null, _t('visual')));

    var activeIdx = 0;

    var viewerContainer = el('div');
    viewerContainer.style.display = 'flex';
    viewerContainer.style.flexDirection = 'column';
    viewerContainer.style.gap = '1.2rem';
    viewerContainer.style.alignItems = 'center';

    var switcher = el('div');
    switcher.style.display = 'flex';
    switcher.style.gap = '.4rem';
    switcher.style.flexWrap = 'wrap';
    switcher.style.marginBottom = '.5rem';

    var imgEl = el('img');
    imgEl.style.maxWidth = '100%';
    imgEl.style.border = '1px solid var(--line)';
    imgEl.style.borderRadius = '8px';
    imgEl.style.boxShadow = '0 6px 16px rgba(11,11,11,.08)';
    imgEl.style.background = '#fff';

    function showPage(idx) {
      activeIdx = idx;
      imgEl.src = images[idx];
      Array.from(switcher.children).forEach(function(child, cIdx) {
        if (cIdx === idx) {
          child.className = 'active';
        } else {
          child.className = '';
        }
      });
    }

    images.forEach(function(src, idx) {
      var pBtn = document.createElement('button');
      pBtn.textContent = _t('page', { n: idx + 1 });
      pBtn.style.padding = '.35rem .7rem';
      pBtn.style.border = '1px solid var(--line)';
      pBtn.style.borderRadius = '6px';
      pBtn.style.background = '#fff';
      pBtn.style.cursor = 'pointer';
      pBtn.style.font = 'inherit';
      pBtn.onclick = function() { showPage(idx); };
      switcher.appendChild(pBtn);
    });

    viewerContainer.appendChild(switcher);
    viewerContainer.appendChild(imgEl);
    container.appendChild(viewerContainer);

    var styleNode = document.createElement('style');
    styleNode.innerHTML = '#lin-screenshots button.active { background: var(--ink) !important; color: #fff !important; }';
    document.head.appendChild(styleNode);

    var screenshotsButton = document.getElementById('lin-tab-screenshots');
    var screenshotsMain = document.getElementById('lin-screenshots-main');

    screenshotsButton.onclick = function () {
      switchTab(screenshotsButton, screenshotsMain);
      showPage(activeIdx);
    };
  }

  function setupTabs() {
    var overview = GRAPH.meta.workbookDoc;
    var ctx = GRAPH.meta.workbookContext;

    var graphButton = document.getElementById('lin-tab-graph');
    var overviewButton = document.getElementById('lin-tab-overview');
    var sheetsButton = document.getElementById('lin-tab-sheets');
    var screenshotsButton = document.getElementById('lin-tab-screenshots');

    var graphMain = document.getElementById('lin-graph-main');
    var overviewMain = document.getElementById('lin-overview-main');
    var sheetsMain = document.getElementById('lin-sheets-main');
    var screenshotsMain = document.getElementById('lin-screenshots-main');

    var overviewTarget = document.getElementById('lin-overview');

    if (overview) {
      overviewButton.hidden = false;
      overviewTarget.innerHTML = '';

      var banner = el('div', 'lin-ai-box');
      var badge = el('span', 'lin-ai-badge', _t('ai_overview'));
      banner.appendChild(badge);
      var desc = el('p', 'lin-hint', _t('ai_overview_desc'));
      banner.appendChild(desc);
      overviewTarget.appendChild(banner);

      var documentBody = el('div', 'lin-doc');
      documentBody.innerHTML = _md(overview);
      overviewTarget.appendChild(documentBody);
    }

    if (ctx && ctx.sheets && ctx.sheets.length) {
      sheetsButton.hidden = false;

      var sidebar = document.getElementById('lin-sheets-sidebar');
      var detailsContainer = document.getElementById('lin-sheet-details');
      sidebar.innerHTML = '';
      detailsContainer.innerHTML = '';

      function showSheetDetails(sheet) {
        detailsContainer.innerHTML = '';

        var card = el('div');
        card.style.background = '#fff';

        var header = el('div');
        header.style.display = 'flex';
        header.style.justifyContent = 'space-between';
        header.style.alignItems = 'center';
        header.style.flexWrap = 'wrap';
        header.style.gap = '.5rem';
        header.style.borderBottom = '1px solid var(--hair)';
        header.style.paddingBottom = '.6rem';
        header.style.marginBottom = '.8rem';

        var title = el('h3', null, '📄 ' + sheet.name);
        title.style.margin = '0';
        title.style.fontSize = '1.25rem';
        title.style.color = 'var(--ink)';

        var dims = el('span', 'lin-hint', sheet.dimensions.rows + ' rows × ' + sheet.dimensions.columns + ' cols');
        dims.style.fontSize = '.85rem';

        header.appendChild(title);
        header.appendChild(dims);
        card.appendChild(header);

        var metaRow = el('div');
        metaRow.style.display = 'flex';
        metaRow.style.gap = '.4rem';
        metaRow.style.flexWrap = 'wrap';
        metaRow.style.marginBottom = '1rem';

        if (sheet.visibility !== 'visible') {
          var visBadge = el('span', 'lin-badge', '👁 ' + sheet.visibility);
          visBadge.style.background = '#e5dbff'; visBadge.style.color = '#5f3dc4';
          metaRow.appendChild(visBadge);
        }
        if (sheet.freeze_panes) {
          var fBadge = el('span', 'lin-badge', '❄ Freeze: ' + sheet.freeze_panes);
          fBadge.style.background = '#e3faf2'; fBadge.style.color = '#0ca678';
          metaRow.appendChild(fBadge);
        }
        if (sheet.hidden_columns && sheet.hidden_columns.length) {
          var hBadge = el('span', 'lin-badge', '🚫 Hidden cols: ' + sheet.hidden_columns.join(', '));
          hBadge.style.background = '#fff0f6'; hBadge.style.color = '#d6336c';
          metaRow.appendChild(hBadge);
        }
        if (sheet.merged_ranges && sheet.merged_ranges.length) {
          var mBadge = el('span', 'lin-badge', '🔗 Merged: ' + sheet.merged_ranges.length);
          mBadge.style.background = '#e8f7ff'; mBadge.style.color = '#0066cc';
          metaRow.appendChild(mBadge);
        }
        if (metaRow.children.length > 0) {
          card.appendChild(metaRow);
        }

        var comments = sheet.comments || [];
        if (comments.length > 0) {
          var commTitle = el('h4', null, '💬 Comments (' + comments.length + ')');
          commTitle.style.fontSize = '.9rem';
          commTitle.style.margin = '1rem 0 .5rem';
          commTitle.style.color = 'var(--ink)';
          card.appendChild(commTitle);

          var commList = el('div');
          commList.style.display = 'flex';
          commList.style.flexDirection = 'column';
          commList.style.gap = '.5rem';
          comments.forEach(function(c) {
            var cBox = el('div');
            cBox.style.background = '#fcfcfb';
            cBox.style.border = '1px solid var(--hair)';
            cBox.style.borderRadius = '6px';
            cBox.style.padding = '.6rem .8rem';
            cBox.style.fontSize = '.85rem';

            var cHead = el('div');
            cHead.style.fontWeight = 'bold';
            cHead.style.marginBottom = '.25rem';
            cHead.style.color = 'var(--ink)';
            cHead.appendChild(document.createTextNode(c.cell + ' (' + c.author + ')'));

            var cText = el('div');
            cText.style.color = 'var(--ink2)';
            cText.appendChild(document.createTextNode(c.text));

            cBox.appendChild(cHead);
            cBox.appendChild(cText);
            commList.appendChild(cBox);
          });
          card.appendChild(commList);
        }

        var screens = GRAPH.meta.screenshots;
        if (screens && typeof screens === 'object' && !Array.isArray(screens) && screens[sheet.name] && screens[sheet.name].length) {
          var sheetImgs = screens[sheet.name];
          var imgTitle = el('h4', null, '📸 ' + _t('visual'));
          imgTitle.style.fontSize = '.9rem';
          imgTitle.style.margin = '1.5rem 0 .5rem';
          imgTitle.style.color = 'var(--ink)';
          card.appendChild(imgTitle);

          var imgGallery = el('div');
          imgGallery.style.display = 'flex';
          imgGallery.style.flexDirection = 'column';
          imgGallery.style.gap = '1rem';
          imgGallery.style.alignItems = 'center';
          imgGallery.style.marginTop = '.5rem';

          sheetImgs.forEach(function(imgSrc) {
            var img = el('img');
            img.src = imgSrc;
            img.style.maxWidth = '100%';
            img.style.border = '1px solid var(--line)';
            img.style.borderRadius = '6px';
            img.style.boxShadow = '0 3px 12px rgba(11,11,11,.05)';
            img.style.background = '#fff';
            imgGallery.appendChild(img);
          });
          card.appendChild(imgGallery);
        }

        detailsContainer.appendChild(card);
      }

      ctx.sheets.forEach(function(sheet, idx) {
        var sBtn = document.createElement('button');
        sBtn.textContent = '📄 ' + sheet.name;
        sBtn.onclick = function() {
          Array.from(sidebar.children).forEach(function(child) { child.className = ''; });
          sBtn.className = 'active';
          showSheetDetails(sheet);
        };
        sidebar.appendChild(sBtn);
        if (idx === 0) {
          sBtn.className = 'active';
          showSheetDetails(sheet);
        }
      });
    }

    graphButton.onclick = function() { switchTab(graphButton, graphMain); };
    overviewButton.onclick = function() { switchTab(overviewButton, overviewMain); };
    sheetsButton.onclick = function() { switchTab(sheetsButton, sheetsMain); };
    screenshotsButton.onclick = function() { switchTab(screenshotsButton, screenshotsMain); };
  }

  function setActive(on, off) { on.classList.add('active'); off.classList.remove('active'); }

  function nodeSize(n) {
    if (n.kind === 'group') return Math.min(64, 30 + 6 * Math.log2((n.count || 1)));
    if (n.kind === 'misc') return 44;
    if (n.kind === 'vba') return 34;
    return 26;
  }
  function shortLabel(n) {
    if (n.kind === 'name') return '📛 ' + n.label;
    if (n.kind === 'vba') return '⚙ ' + n.label;
    return n.label;
  }
  function layoutOpts(name, hasFcose) {
    if (name === 'dagre') return { name: 'dagre', rankDir: 'LR', nodeSep: 26,
      rankSep: 90, edgeSep: 12, animate: false };
    if (name === 'fcose' && hasFcose) return { name: 'fcose', quality: 'default',
      animate: false, nodeRepulsion: 5500, idealEdgeLength: 90, packComponents: true };
    return { name: 'cose', animate: false };
  }
  function buildStyle(big) {
    var s = [
      { selector: 'node', style: {
        label: 'data(label)', width: 'data(size)', height: 'data(size)',
        'font-size': 9, 'min-zoomed-font-size': 8, 'text-valign': 'bottom',
        'text-margin-y': 4, 'text-wrap': 'ellipsis', 'text-max-width': 130,
        color: '#52514e', 'border-width': 1.5, 'border-color': 'rgba(11,11,11,0.18)' } },
      { selector: 'node[degree <= 1]', style: { 'opacity': 0.55 } },
      { selector: 'node[degree = 2]', style: { 'opacity': 0.75 } },
      { selector: 'edge', style: {
        width: 1.4, 'curve-style': big ? 'straight' : 'bezier',
        'target-arrow-shape': 'triangle', 'arrow-scale': 0.75,
        'line-color': '#c3c2b7', 'target-arrow-color': '#c3c2b7' } },
      { selector: 'edge.name', style: { 'line-style': 'dashed', 'line-color': '#d99a19',
        'target-arrow-color': '#d99a19' } },
      { selector: 'edge.call', style: { 'line-style': 'dotted', 'line-color': '#4a3aa7',
        'target-arrow-color': '#4a3aa7' } },
      { selector: 'edge.vba-write', style: { 'line-color': '#eb6834',
        'target-arrow-color': '#eb6834' } },
      { selector: 'edge.vba-read', style: { 'line-style': 'dashed', 'line-color': '#eb6834',
        'target-arrow-color': '#eb6834' } },
      { selector: 'edge.approx', style: { opacity: 0.55 } },
      { selector: 'node:selected', style: { 'border-width': 3.5,
        'border-color': '#0b0b0b', color: '#0b0b0b', 'font-size': 11 } },
      { selector: '.dimmed', style: { opacity: 0.12 } },
      { selector: 'edge.hl', style: { width: 2.6, opacity: 1 } }
    ];
    Object.keys(KIND).forEach(function (k) {
      s.push({ selector: 'node.' + k,
        style: { 'background-color': KIND[k].color, shape: KIND[k].shape } });
    });
    return s;
  }
  function buildLegend() {
    var present = {};
    GRAPH.nodes.forEach(function (n) { present[n.kind] = true; });
    var el = document.getElementById('lin-legend');
    Object.keys(KIND).forEach(function (k) {
      if (!present[k]) return;
      var span = document.createElement('span');
      var sw = document.createElement('span');
      sw.className = 'lin-sw'; sw.style.background = KIND[k].color;
      span.appendChild(sw);
      span.appendChild(document.createTextNode(_t(KIND[k].labelKey)));
      el.appendChild(span);
    });
  }

  function fmt(v) {
    if (v === null || v === undefined) return '—';
    if (typeof v === 'number') return Number.isInteger(v) ? String(v)
      : v.toLocaleString(LANG, { maximumFractionDigits: 4 });
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'object') {
      if (v.range) return v.range + ' (' + v.n + ' ' + _t('cells') + ')';
      if (v.type === 'Error' && ERRTEXT[v.kind]) return ERRTEXT[v.kind];
      return JSON.stringify(v);
    }
    return String(v);
  }

  // Narrow the graph to one sheet. Nodes of another sheet are kept — dimmed —
  // when an edge ties them to the selected one, because a formula that reads
  // from elsewhere is only readable if that "elsewhere" is on screen.
  function applySheetFilter(cy, sheet, layoutName, hasFcose) {
    var before = cy.nodes(':visible').length;
    CUR_SHEET = sheet || ALL_SHEETS;
    cy.batch(function () {
      cy.elements().removeClass('dimmed hl');
      if (CUR_SHEET === ALL_SHEETS) { cy.elements().show(); return; }
      var inSheet = cy.nodes().filter(function (n) {
        return n.data('sheet') === CUR_SHEET;
      });
      var keep = inSheet.union(inSheet.neighborhood().nodes());
      var keepEdges = cy.edges().filter(function (e) {
        return keep.contains(e.source()) && keep.contains(e.target());
      });
      cy.elements().hide();
      keep.show(); keepEdges.show();
    });
    clearSel(cy);  // the previously selected node may now be hidden
    var visible = cy.elements(':visible');
    visible.layout(layoutOpts(layoutName, hasFcose)).run();
    if (cy.nodes(':visible').length !== before) cy.fit(visible, 40);
  }
  // Cross-sheet neighbours read as "external": dimmed, but present. Reapplied
  // after every selection change, which clears the class wholesale.
  function dimExternal(cy) {
    if (CUR_SHEET === ALL_SHEETS) return;
    cy.nodes(':visible').filter(function (n) {
      return n.data('sheet') !== CUR_SHEET;
    }).addClass('dimmed');
  }

  function select(cy, id) {
    var n = byId[id]; if (!n) return;
    var ele = cy.getElementById(id);
    cy.elements().removeClass('dimmed hl');
    var hood = ele.closedNeighborhood();
    cy.elements().not(hood).addClass('dimmed');
    hood.edges().addClass('hl');
    dimExternal(cy);
    cy.nodes(':selected').unselect(); ele.select();
    renderPanel(cy, n);
  }
  function clearSel(cy) {
    cy.elements().removeClass('dimmed hl'); cy.nodes(':selected').unselect();
    dimExternal(cy);
    renderPlaceholder();
  }

  function renderPlaceholder() {
    var p = document.getElementById('lin-panel');
    p.classList.remove('hidden'); p.innerHTML = '';

    var icon = el('div');
    icon.style.textAlign = 'center';
    icon.style.fontSize = '2.5rem';
    icon.style.marginTop = '4rem';
    icon.style.color = 'var(--muted)';
    icon.textContent = '🔍';
    p.appendChild(icon);

    var title = el('h2', null, _t('placeholder_title'));
    title.style.textAlign = 'center';
    title.style.marginTop = '1rem';
    p.appendChild(title);

    var desc = el('p', 'lin-hint');
    desc.style.textAlign = 'center';
    desc.style.lineHeight = '1.45';
    desc.style.padding = '0 .5rem';
    desc.textContent = _t('placeholder_desc');
    p.appendChild(desc);
  }

  function el(tag, cls, txt) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (txt !== undefined) e.textContent = txt;
    return e;
  }
  function section(title) { var s = el('div'); s.appendChild(el('h3', null, title)); return s; }

  // Mini-markdown → HTML. Handles **bold**, *italic*, `code`, lists -, headings ###.
  // Enough for the AI cards, no external dep. Escapes & < > first, so model
  // output can never inject markup.
  function _md(src) {
    var esc = src.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    var lines = esc.split('\n'), out = [], inList = false;
    for (var i = 0; i < lines.length; i++) {
      var l = lines[i];
      if (/^###\s/.test(l)) { if (inList) { out.push('</ul>'); inList = false; } out.push('<h3>' + l.slice(4).trim() + '</h3>'); continue; }
      if (/^##\s/.test(l))  { if (inList) { out.push('</ul>'); inList = false; } out.push('<h2>' + l.slice(3).trim() + '</h2>'); continue; }
      if (/^#\s/.test(l))   { if (inList) { out.push('</ul>'); inList = false; } out.push('<h1>' + l.slice(2).trim() + '</h1>'); continue; }
      if (/^[-*]\s/.test(l)) { if (!inList) { out.push('<ul>'); inList = true; } out.push('<li>' + _mdInline(l.replace(/^[-*]\s+/, '')) + '</li>'); continue; }
      if (inList) { out.push('</ul>'); inList = false; }
      out.push('<p>' + _mdInline(l) + '</p>');
    }
    if (inList) out.push('</ul>');
    return out.join('');
  }
  function _mdInline(s) {
    return s
      .replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      .replace(/\*([^*]+)\*/g, '<em>$1</em>');
  }

  function renderPanel(cy, n) {
    var p = document.getElementById('lin-panel');
    p.classList.remove('hidden'); p.innerHTML = '';
    var head = el('div'); head.style.display = 'flex'; head.style.alignItems = 'center';
    var badgeLabel = (KIND[n.kind] || {}).labelKey ? _t(KIND[n.kind].labelKey) : n.kind;
    var badge = el('span', 'lin-badge', badgeLabel);
    badge.style.background = (KIND[n.kind] || {}).color || '#898781';
    head.appendChild(badge);
    var close = el('button', 'lin-close', '✕'); close.onclick = function () { clearSel(cy); };
    head.appendChild(close); p.appendChild(head);
    p.appendChild(el('h2', null, n.label));

    if (n.formula) {
      var sf = section(_t('formula')); sf.appendChild(el('code', 'lin-formula', n.formula));
      if (n.kind === 'group') {
        sf.appendChild(el('p', 'lin-hint',
          _t('stretched_pattern', { count: n.count.toLocaleString(LANG), bbox: n.bbox })));
        sf.appendChild(el('code', 'lin-formula', n.r1c1));
      }
      p.appendChild(sf);
      var sv = valueSection(n);
      if (n.samples && n.samples.length) {
        n.samples.forEach(function (s) {
          var line = s.addr + ' = ' + fmt(s.value);
          if (s.date) line += ' (' + s.date + ')';
          if (SRC[s.source] && s.source !== 'engine') line += ' [' + _t(SRC[s.source].labelKey) + ']';
          sv.appendChild(el('div', 'lin-hint', line));
        });
      }
      p.appendChild(sv);
    }
    if (n.kind === 'input' && n.values && n.values.length) {
      var si = section(_t('value_samples'));
      n.values.forEach(function (s) {
        si.appendChild(el('div', 'lin-hint', s.addr + ' = ' + fmt(s.value)));
      });
      if (n.count > n.values.length)
        si.appendChild(el('div', 'lin-hint', '… ' + n.count.toLocaleString(LANG) + ' ' + _t('cells')));
      p.appendChild(si);
    }
    if (n.kind === 'name' && n.targets) {
      var sn = section(_t('target')); sn.appendChild(el('code', 'lin-formula', n.targets.join(' ; ')));
      p.appendChild(sn);
      if (n.value !== undefined && n.value !== null) p.appendChild(valueSection(n));
    }
    if (n.steps && (n.steps.children && n.steps.children.length || (n.steps.inputs && n.steps.inputs.length))) {
      var ss = section(_t('step_decomp'));
      ss.appendChild(el('p', 'lin-hint', _t('step_hint')));
      renderStep(ss, n.steps, 0);
      p.appendChild(ss);
    }
    if (n.kind === 'vba') {
      var sc = section(n.procKind + ' — module ' + n.module);
      sc.appendChild(el('pre', 'lin-code', n.code || '')); p.appendChild(sc);
    }
    if (n.doc) {
      var sd = el('div', 'lin-ai-box');
      var badge = el('span', 'lin-ai-badge', _t('ai_doc'));
      sd.appendChild(badge);
      var dv = el('div', 'lin-doc');
      dv.innerHTML = _md(n.doc);
      sd.appendChild(dv); p.appendChild(sd);
    }
    appendNav(cy, p, _t('precedents'), GRAPH.edges.filter(function (e) { return e.target === n.id; })
      .map(function (e) { return { node: byId[e.source], kind: e.kind }; }));
    appendNav(cy, p, _t('dependents'), GRAPH.edges.filter(function (e) { return e.source === n.id; })
      .map(function (e) { return { node: byId[e.target], kind: e.kind }; }));
  }

  // "Computed value" section: the value (or its date), a provenance badge, and
  // — when the engine disagrees with what the file stored — both figures.
  function valueSection(n) {
    var sv = section(_t('computed_value'));
    var shown = n.valueDate || fmt(n.value);
    var line = el('div', 'lin-val', shown);
    var src = SRC[n.valueSource];
    if (src) {
      var badge = el('span', 'lin-src', _t(src.labelKey));
      badge.style.background = src.color;
      line.appendChild(badge);
    }
    sv.appendChild(line);
    if (n.cachedValue !== undefined && n.cachedValue !== null) {
      var cached = fmt(n.cachedValue);
      // a date node compares on the date part only: a file-cached datetime
      // string ("2026-08-07 00:00:00") is the same day as valueDate
      var sameDate = n.valueDate && typeof n.cachedValue === 'string' &&
        n.cachedValue.slice(0, 10) === n.valueDate;
      if (!sameDate && cached !== shown) {
        sv.appendChild(el('div', 'lin-cmp', _t('value_cached') + ': ' + cached));
        sv.appendChild(el('div', 'lin-cmp', _t('differs_from_file', { recalc: shown })));
      }
    }
    return sv;
  }

  function renderStep(parent, s, depth) {
    var d = el('div', 'lin-step' + (depth === 0 ? ' lin-final' : ''));
    d.style.marginLeft = Math.min(depth, 6) * 12 + 'px';
    if (depth === 0) d.appendChild(el('div', 'lin-val', _t('final_result')));
    var head = el('div');
    head.appendChild(el('span', 'lin-op', s.label));
    var expr = el('span', 'lin-expr', s.expr); head.appendChild(expr);
    d.appendChild(head);
    var val = el('div');
    if (s.evaluated) { val.appendChild(document.createTextNode('= '));
      var b = el('b', null, fmt(s.value)); val.appendChild(b); }
    else val.appendChild(el('span', 'lin-hint', _t('not_evaluated')));
    d.appendChild(val);
    (s.inputs || []).forEach(function (inp) {
      if (inp.ref !== undefined) d.appendChild(el('span', 'lin-in', inp.ref + ' = ' + fmt(inp.value)));
      else d.appendChild(el('span', 'lin-in lit', fmt(inp.literal)));
    });
    parent.appendChild(d);
    (s.children || []).forEach(function (c) { renderStep(parent, c, depth + 1); });
  }

  function appendNav(cy, parent, title, items) {
    items = items.filter(function (x) { return x.node; });
    if (!items.length) return;
    var sec = section(title + ' (' + items.length + ')');
    var ul = el('ul', 'lin-nav');
    items.forEach(function (it) {
      var li = el('li');
      var btn = el('button');
      var dot = el('span', 'lin-dot'); dot.style.background = (KIND[it.node.kind] || {}).color || '#898781';
      btn.appendChild(dot); btn.appendChild(document.createTextNode(it.node.label));
      btn.onclick = (function (id) { return function () {
        select(cy, id);
        cy.animate({ center: { eles: cy.getElementById(id) }, duration: 250 });
      }; })(it.node.id);
      li.appendChild(btn); li.appendChild(el('span', 'lin-ek', it.kind));
      ul.appendChild(li);
    });
    sec.appendChild(ul); parent.appendChild(sec);
  }

  if (document.readyState === 'loading')
    document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
</script>
"""
