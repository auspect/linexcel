# Offline dashboard graph

The dashboard keeps its existing HTML/JavaScript application and mounts one React
island in `#graph-canvas`. `src/graph.jsx` owns the React Flow tree, custom cell,
formula-group and sheet cards, library controls, minimap, and Dagre layout. The
shell calls the small `LinexcelGraph.mount()` interface. It retains imports,
history, captures, AI controls, Datasette, exports, accessible relation lists and
French/English translation. The separate standalone HTML report still uses its
existing Cytoscape renderer and bundled assets.

Python users need no Node installation or network access. The committed
`src/linexcel/assets/app-graph.js` and `.css` are served from an explicit allowlist
under the existing mounted, authenticated application and CSP. The IIFE includes
React, React DOM, React Flow and Dagre; it has no dynamic imports or CDN requests.
Package versions are exact in `package.json`; transitive versions are locked.
The runtime libraries use MIT and BSD-3-Clause licenses; their complete notices
are in `THIRD_PARTY_NOTICES.md`, also packaged in the wheel. Vite is build-only.

With Node 24 and npm 11:

```sh
cd frontend
npm ci
npm run build
npm run check
```

`check` rebuilds in memory and compares every byte of the committed JS and CSS.
CI runs it. When changing library versions, update all bundled runtime library
licenses in `THIRD_PARTY_NOTICES.md`. Rebuild and commit both assets with sources.

The existing `POST /api/projects/:id/tasks` operation `build_graph` produces a
versioned, immutable-index-specific SQLite cache in an isolated budgeted worker.
The graph tab starts that operation automatically, sharing/caching server tasks.
`GET /api/projects/:id/graph-view` returns the root, `?view=:id` drills into a
group, `?sheet=:name` finds a sheet, and `?nodeId=:id` locates a cell's leaf group.
Views contain at most 32 children plus an explicit rest-of-workbook card. Both
cross-group edge weights and internal edge counts are SQL aggregates; ranges
retain their compact index semantics. R1C1 pattern membership comes from the
index, not from displayed formula text. A sheet or pattern above 32 children has
additional bounded partition levels. Search continues to use disk-index pages.

Formula selection waits 180 ms before submitting a targeted calculation. It
aborts superseded neighborhood reads, cancels known superseded automatic tasks,
reuses completed results for the immutable project revision, and leaves explicit
Recalculate as the force-refresh path. Selecting inputs or static VBA never
evaluates them. Server time, memory, closure and output limits still apply.
Deep links use `?project=:id&node=Sheet!A1&tab=nodes` (or `tab=graph`).

For a real-browser dashboard capture run, start a dedicated local server on 8792
using `validation_screenshots/frontend-server` as its data directory, then run:

```sh
uv run --extra screenshots python scripts/validate_frontend.py --workbook /path/to/synthetic.xlsx
```

This records every tab at 1440×900, 2560×1440 and 390×844, graph hierarchy levels,
automatic decomposition, browser errors, horizontal overflow, and visible notice
scoping. Screenshots and manifests stay ignored. This supplements, rather than
replaces, the full `AGENTS.md` acceptance workflow and independent visual review.
