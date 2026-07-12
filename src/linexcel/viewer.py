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
from functools import lru_cache
from pathlib import Path
from typing import Any

_ASSETS = Path(__file__).parent / "assets"

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
    graph: dict[str, Any], title: str = "Lineage Excel", full_document: bool = True
) -> str:
    """Build the viewer HTML for a given graph."""
    data = _safe_json(graph)
    body = _TEMPLATE.replace("__GRAPH_JSON__", data).replace(
        "__TITLE__", _escape_text(title)
    )
    if not full_document:
        return body
    scripts = _inline_scripts()
    if scripts is None:
        scripts = "\n".join(f'<script src="{url}"></script>' for url in _CDN)
    return (
        "<!doctype html><html lang='en'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width, initial-scale=1'>"
        f"<title>{_escape_text(title)}</title>{scripts}</head>"
        f"<body>{body}</body></html>"
    )


def _escape_text(text: str) -> str:
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
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
  .lin-step {
    border-left: 3px solid var(--blue); background: #fff; border-radius: 0 6px 6px 0;
    padding: .3rem .5rem; margin: .35rem 0; box-shadow: 0 1px 2px rgba(11,11,11,.05);
  }
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
</style>

<div class="lin-root">
  <div class="lin-bar">
    <h1>__TITLE__</h1>
    <span class="stat" id="lin-stats"></span>
    <button id="lin-tab-graph" class="lin-tab active">Graph</button>
    <button id="lin-tab-overview" class="lin-tab" hidden>Workbook overview</button>
    <span style="flex:1"></span>
    <input id="lin-search" placeholder="Search… ⏎" />
    <button id="lin-lay-dagre" class="active">Flow</button>
    <button id="lin-lay-fcose">Organic</button>
    <button id="lin-fit">Fit</button>
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
</div>

<script>
(function () {
  var GRAPH = __GRAPH_JSON__;
  var KIND = {
    cell:  { color: '#2a78d6', shape: 'round-rectangle', label: 'Formula' },
    group: { color: '#4a3aa7', shape: 'round-rectangle', label: 'Stretched formulas' },
    input: { color: '#1baf7a', shape: 'ellipse', label: 'Source data' },
    name:  { color: '#eda100', shape: 'diamond', label: 'Defined name' },
    vba:   { color: '#eb6834', shape: 'hexagon', label: 'VBA' },
    misc:  { color: '#898781', shape: 'octagon', label: 'Other (aggregated)' },
    opaque:{ color: '#898781', shape: 'ellipse', label: 'External reference' }
  };
  var byId = {};
  GRAPH.nodes.forEach(function (n) { byId[n.id] = n; });

  function boot() {
    setupTabs();
    var cyContainer = document.getElementById('lin-cy');
    if (typeof cytoscape === 'undefined') {
      var f = document.createElement('div');
      f.className = 'lin-fallback';
      f.textContent = 'Cytoscape could not be loaded (CDN access required). '
        + 'The JSON graph remains available via result.to_dict().';
      cyContainer.appendChild(f);
      return;
    }
    // Register layout extensions if present.
    var hasFcose = false;
    try {
      if (window.cytoscapeFcose) { cytoscape.use(window.cytoscapeFcose); hasFcose = true; }
      if (window.cytoscapeDagre) { cytoscape.use(window.cytoscapeDagre); }
    } catch (e) { /* already registered */ }

    var stats = GRAPH.meta.stats;
    document.getElementById('lin-stats').textContent =
      stats.totalFormulas.toLocaleString('en') + ' formulas · ' +
      stats.totalNodes + ' nodes · ' + stats.totalEdges + ' edges' +
      (stats.vbaProcs ? ' · ' + stats.vbaProcs + ' VBA' : '');

    var big = GRAPH.nodes.length + GRAPH.edges.length > 2500;
    var elements = [];
    GRAPH.nodes.forEach(function (n) {
      elements.push({ data: {
        id: n.id, label: shortLabel(n), size: nodeSize(n)
      }, classes: n.kind });
    });
    GRAPH.edges.forEach(function (e) {
      elements.push({ data: { id: e.id, source: e.source, target: e.target },
        classes: e.kind + (e.approx ? ' approx' : '') });
    });

    var hasDagre = typeof dagre !== 'undefined' || window.cytoscapeDagre;
    var initial = hasDagre ? 'dagre' : (hasFcose ? 'fcose' : 'cose');
    var cy = cytoscape({
      container: cyContainer, elements: elements,
      minZoom: 0.05, maxZoom: 4, wheelSensitivity: 0.25,
      pixelRatio: big ? 1 : 'auto', textureOnViewport: big, hideEdgesOnViewport: big,
      style: buildStyle(big), layout: layoutOpts(initial, hasFcose)
    });

    cy.on('tap', 'node', function (ev) { select(cy, ev.target.id()); });
    cy.on('tap', function (ev) { if (ev.target === cy) clearSel(cy); });

    document.getElementById('lin-fit').onclick = function () {
      cy.animate({ fit: { padding: 40 }, duration: 250 });
    };
    var bd = document.getElementById('lin-lay-dagre');
    var bf = document.getElementById('lin-lay-fcose');
    bd.onclick = function () {
      setActive(bd, bf); cy.layout(layoutOpts('dagre', hasFcose)).run();
    };
    bf.onclick = function () {
      setActive(bf, bd); cy.layout(layoutOpts('fcose', hasFcose)).run();
    };
    if (initial !== 'dagre') setActive(bf, bd);
    var search = document.getElementById('lin-search');
    search.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter') return;
      var q = search.value.trim().toLowerCase();
      if (!q) return;
      var hit = GRAPH.nodes.find(function (n) {
        return (n.label || '').toLowerCase().indexOf(q) >= 0
          || (n.formula || '').toLowerCase().indexOf(q) >= 0;
      });
      if (hit) { select(cy, hit.id);
        cy.animate({ center: { eles: cy.getElementById(hit.id) },
          zoom: Math.max(cy.zoom(), 1), duration: 300 }); }
    });
    buildLegend();
  }

  function setupTabs() {
    var overview = GRAPH.meta.workbookDoc;
    if (!overview) return;
    var graphButton = document.getElementById('lin-tab-graph');
    var overviewButton = document.getElementById('lin-tab-overview');
    var graphMain = document.getElementById('lin-graph-main');
    var overviewMain = document.getElementById('lin-overview-main');
    var target = document.getElementById('lin-overview');
    overviewButton.hidden = false;
    target.appendChild(el('h2', null, 'Workbook overview'));
    var documentBody = el('div', 'lin-doc');
    documentBody.innerHTML = _md(overview);
    target.appendChild(documentBody);
    graphButton.onclick = function () {
      graphButton.classList.add('active'); overviewButton.classList.remove('active');
      graphMain.classList.remove('hidden'); overviewMain.classList.add('hidden');
    };
    overviewButton.onclick = function () {
      overviewButton.classList.add('active'); graphButton.classList.remove('active');
      overviewMain.classList.remove('hidden'); graphMain.classList.add('hidden');
    };
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
      span.appendChild(document.createTextNode(KIND[k].label));
      el.appendChild(span);
    });
  }

  function fmt(v) {
    if (v === null || v === undefined) return '—';
    if (typeof v === 'number') return Number.isInteger(v) ? String(v)
      : v.toLocaleString('en', { maximumFractionDigits: 4 });
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'object') return v.range ? (v.range + ' (' + v.n + ' cells)')
      : JSON.stringify(v);
    return String(v);
  }

  function select(cy, id) {
    var n = byId[id]; if (!n) return;
    var ele = cy.getElementById(id);
    cy.elements().removeClass('dimmed hl');
    var hood = ele.closedNeighborhood();
    cy.elements().not(hood).addClass('dimmed');
    hood.edges().addClass('hl');
    cy.nodes(':selected').unselect(); ele.select();
    renderPanel(cy, n);
  }
  function clearSel(cy) {
    cy.elements().removeClass('dimmed hl'); cy.nodes(':selected').unselect();
    document.getElementById('lin-panel').classList.add('hidden');
  }

  function el(tag, cls, txt) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (txt !== undefined) e.textContent = txt;
    return e;
  }
  function section(title) { var s = el('div'); s.appendChild(el('h3', null, title)); return s; }

  // ponytail: mini-markdown → HTML. Handles **bold**, *italic*, `code`, lists -, headings ###.
  // Enough for Gemini cards, no external dep.
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
    var badge = el('span', 'lin-badge', (KIND[n.kind] || {}).label || n.kind);
    badge.style.background = (KIND[n.kind] || {}).color || '#898781';
    head.appendChild(badge);
    var close = el('button', 'lin-close', '✕'); close.onclick = function () { clearSel(cy); };
    head.appendChild(close); p.appendChild(head);
    p.appendChild(el('h2', null, n.label));

    if (n.formula) {
      var sf = section('Formula'); sf.appendChild(el('code', 'lin-formula', n.formula));
      if (n.kind === 'group') {
        sf.appendChild(el('p', 'lin-hint',
          'Stretched pattern over ' + n.count.toLocaleString('en') + ' cells (' + n.bbox + ').'));
        sf.appendChild(el('code', 'lin-formula', n.r1c1));
      }
      p.appendChild(sf);
      var sv = section('Computed value'); sv.appendChild(el('div', 'lin-val', fmt(n.value)));
      if (n.samples && n.samples.length) {
        n.samples.forEach(function (s) {
          sv.appendChild(el('div', 'lin-hint', s.addr + ' = ' + fmt(s.value)));
        });
      }
      p.appendChild(sv);
    }
    if (n.kind === 'input' && n.values && n.values.length) {
      var si = section('Value samples');
      n.values.forEach(function (s) {
        si.appendChild(el('div', 'lin-hint', s.addr + ' = ' + fmt(s.value)));
      });
      if (n.count > n.values.length)
        si.appendChild(el('div', 'lin-hint', '… ' + n.count.toLocaleString('en') + ' cells'));
      p.appendChild(si);
    }
    if (n.kind === 'name' && n.targets) {
      var sn = section('Target'); sn.appendChild(el('code', 'lin-formula', n.targets.join(' ; ')));
      p.appendChild(sn);
    }
    if (n.steps && (n.steps.children && n.steps.children.length || (n.steps.inputs && n.steps.inputs.length))) {
      var ss = section('Step-by-step decomposition');
      ss.appendChild(el('p', 'lin-hint', 'Each function/operator is evaluated individually.'));
      renderStep(ss, n.steps, 0);
      p.appendChild(ss);
    }
    if (n.kind === 'vba') {
      var sc = section(n.procKind + ' — module ' + n.module);
      sc.appendChild(el('pre', 'lin-code', n.code || '')); p.appendChild(sc);
    }
    // ponytail: mini-markdown renderer for AI docs (no external dep)
    if (n.doc) {
      var sd = section('AI Documentation');
      var dv = el('div', 'lin-doc');
      dv.innerHTML = _md(n.doc);
      sd.appendChild(dv); p.appendChild(sd);
    }
    appendNav(cy, p, 'Precedents', GRAPH.edges.filter(function (e) { return e.target === n.id; })
      .map(function (e) { return { node: byId[e.source], kind: e.kind }; }));
    appendNav(cy, p, 'Dependents', GRAPH.edges.filter(function (e) { return e.source === n.id; })
      .map(function (e) { return { node: byId[e.target], kind: e.kind }; }));
  }

  function renderStep(parent, s, depth) {
    var d = el('div', 'lin-step');
    d.style.marginLeft = Math.min(depth, 6) * 12 + 'px';
    var head = el('div');
    head.appendChild(el('span', 'lin-op', s.label));
    var expr = el('span', 'lin-expr', s.expr); head.appendChild(expr);
    d.appendChild(head);
    var val = el('div');
    if (s.evaluated) { val.appendChild(document.createTextNode('= '));
      var b = el('b', null, fmt(s.value)); val.appendChild(b); }
    else val.appendChild(el('span', 'lin-hint', 'not evaluated'));
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
