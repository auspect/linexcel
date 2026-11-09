import React, { memo, useEffect, useState, useRef } from "react";
import { createRoot } from "react-dom/client";
import { flushSync } from "react-dom";
import {
  ReactFlow,
  ReactFlowProvider,
  MiniMap,
  Controls,
  Background,
  Handle,
  Position,
  MarkerType,
  applyNodeChanges,
  Panel,
  BaseEdge,
} from "@xyflow/react";
import dagre from "@dagrejs/dagre";
import "@xyflow/react/dist/style.css";
import "./graph.css";
import "./shell.css";
import "./theme.css";
import "./library-search.css";
export { compressNeighborhood } from './aggregation.js';

const Card = memo(({ data, selected }) => (
  <article
    data-no-i18n=""
    className={`flow-card ${data.tone || ""} ${selected ? "is-selected" : ""}`}
    title={[data.title,data.subtitle].filter(Boolean).join('\n')}
  >
    {Object.entries({left:Position.Left,right:Position.Right,top:Position.Top,bottom:Position.Bottom}).flatMap(([side,position])=>['source','target'].map(type=><Handle key={`${type}-${side}`} id={`${type}-${side}`} type={type} position={position} isConnectable={false} />))}
    <span className="flow-kind">{data.kindLabel}</span>
    <div className="flow-title-row"><strong>{data.title}</strong>{data.count!==undefined&&<span className="flow-count">{data.count}</span>}</div>
    <span className="flow-subtitle">{data.subtitle}</span>
    {data.valueLabel&&<span className="flow-value">{data.valueLabel}</span>}
  </article>
));
const nodeTypes = { cell: Card, formulaGroup: Card, sheet: Card, cellGroup: Card };
function LayeredEdge({
  id,
  sourceX,
  sourceY,
  targetX,
  targetY,
  sourcePosition,
  targetPosition,
  markerEnd,
  data,
  label,
}) {
  const points = (
    data?.points || [
      { x: sourceX, y: sourceY },
      { x: targetX, y: targetY },
    ]
  ).map((p) => ({ ...p }));
  if (!data?.bundle) {
    const boundary=(x,y,side)=>({x:x+(side==='right'?-3:side==='left'?3:0),y:y+(side==='bottom'?-3:side==='top'?3:0)});
    const start=boundary(sourceX,sourceY,sourcePosition),end=boundary(targetX,targetY,targetPosition);
    const first=points[0],last=points.at(-1);
    // Dagre's terminal intersection need not be at the handle's midpoint.
    // Join along the card border; replacing the intersection with the handle
    // made a new diagonal that could cut through an unrelated card.
    points.unshift(start,sourcePosition==='left'||sourcePosition==='right'?{x:first.x,y:start.y}:{x:start.x,y:first.y});
    points.push(targetPosition==='left'||targetPosition==='right'?{x:last.x,y:end.y}:{x:end.x,y:last.y},end);
  }
  const path = points.map((p, i) => `${i ? "L" : "M"}${p.x},${p.y}`).join(" "),
    center = data?.bundle ? {x:(points[0].x+points.at(-1).x)/2,y:(points[0].y+points.at(-1).y)/2} : points[Math.floor(points.length / 2)];
  return (
    <>
      <BaseEdge
        id={id}
        path={path}
        markerEnd={markerEnd}
        label={data?.bundleLabel || label}
        labelX={center.x}
        labelY={center.y}
      />
    </>
  );
}
const edgeTypes = { layered: LayeredEdge };

function Dock({ elements = [] }) {
  const ref = useRef(null);
  useEffect(() => {
    const slots = elements.map((element) => {
      const slot = document.createComment("graph control position");
      element.before(slot);
      ref.current.append(element);
      return { element, slot };
    });
    return () =>
      slots.forEach(({ element, slot }) => {
        slot.after(element);
        slot.remove();
      });
  }, []);
  return <div ref={ref} className="graph-docked-controls" />;
}

function layered(nodes, edges, portrait, aspect, availableWidth, availableHeight) {
  const owned=nodes.filter(n=>n.data.tone!=='outside'),outside=nodes.find(n=>n.data.tone==='outside');
  const ownedIds=new Set(owned.map(n=>n.id));
  const leaf=owned.length>1&&owned.every(n=>n.data.original.nodeId)&&!edges.some(e=>ownedIds.has(e.source)&&ownedIds.has(e.target));
  if(leaf){
    const columns=Math.max(1,Math.min(6,Math.ceil(Math.sqrt(owned.length*aspect/2.23))));
    owned.forEach((n,i)=>{n.position={x:(i%columns)*272,y:Math.floor(i/columns)*144};});
    const width=columns*272-40,height=Math.ceil(owned.length/columns)*144-40;
    if(outside)outside.position={x:width+110,y:Math.max(0,(height-104)/2)};
    // All exact cell relationships remain in the text list/model. Render only
    // one weighted link per direction between the grid boundary and outside.
    // Individual bidirectional buses used to draw a box around every card.
    for (const incoming of [false, true]) {
      const bundle = edges.filter(e => outside && (incoming ? e.source : e.target) === outside.id);
      bundle.forEach(e => {e.hidden = true;});
      if (!bundle.length) continue;
      const edge = bundle[0], y = height/2 + (incoming ? -14 : 14);
      const grid = {x:width+8,y}, rest = {x:outside.position.x,y:outside.position.y+52+(incoming?-14:14)};
      edge.hidden = false;
      edge.sourceHandle = incoming ? 'source-left' : 'source-right';
      edge.targetHandle = incoming ? 'target-right' : 'target-left';
      edge.data = {...edge.data,bundle:true,points:incoming?[rest,grid]:[grid,rest],
        bundleLabel:String(bundle.reduce((sum,e)=>sum+(e.weight||1),0))};
    }
    return nodes;
  }
  const build = (direction, align, ranker='network-simplex', nodesep=20) => {
  const graph = new dagre.graphlib.Graph({ multigraph: true });
  graph.setGraph({
    rankdir: direction,
    ranksep: 70,
    nodesep,
    edgesep: 8,
    align,
    ranker,
    marginx: 30,
    marginy: 30,
  });
  graph.setDefaultEdgeLabel(() => ({}));
  [...nodes].sort((a,b)=>a.id.localeCompare(b.id)).forEach((node) => graph.setNode(node.id, { width: 232, height: 104 }));
  edges.forEach((edge) =>
    graph.setEdge(
      edge.source,
      edge.target,
      { weight: Math.min(Number(edge.label) || 1, 100) },
      edge.id,
    ),
  );
  dagre.layout(graph);
  return graph;
  };
  const fit = g => Math.min(availableWidth/(g.graph().width-60), availableHeight/(g.graph().height-60));
  let graph = build("LR");
  // Dagre offers several legal alignments. Choose a tighter layered variant
  // for a small neighborhood before considering any oversized-rank fallback.
  if(nodes.length<=40 && fit(graph)<.5){
    for(const ranker of ['network-simplex','longest-path'])for(const align of ['UL','DL']){
      const candidate=build('LR',align,ranker,8);
      if(fit(candidate)>fit(graph))graph=candidate;
    }
  }
  if (portrait) {
    const vertical = build("TB");
    portrait = fit(vertical)>fit(graph);
    if (portrait) graph = vertical;
  }
  // Pack unusually broad ranks after Dagre assigns their dependency layers.
  // The bounded overview remains useful even when siblings cannot be grouped.
  const ranks=new Map();
  for(const n of nodes){const p=graph.node(n.id),key=portrait?p.y:p.x;
    if(!ranks.has(key))ranks.set(key,[]);ranks.get(key).push(n);}
  const maxRank=Math.max(...[...ranks.values()].map(rank=>rank.length));
  // Normal neighborhoods retain Dagre's dependency layers. A capped heterogeneous
  // fanout can still exceed the viewport after semantic grouping; only wrap
  // genuinely oversized ranks, inside their own dependency layer.
  if(nodes.length>40 && maxRank>16 && availableHeight/(maxRank*124)<.35){
    let offset=30;
    for(const [,rank] of [...ranks].sort((a,b)=>a[0]-b[0])){
      const wraps=rank.length>16?Math.min(4,Math.ceil(rank.length/Math.max(8,Math.floor(availableHeight/(144*.45))))):1;
      const rows=portrait?wraps:Math.ceil(rank.length/wraps),columns=portrait?Math.ceil(rank.length/wraps):wraps;
      rank.sort((a,b)=>(portrait?graph.node(a.id).x-graph.node(b.id).x:graph.node(a.id).y-graph.node(b.id).y));
      rank.forEach((n,i)=>{const p=graph.node(n.id),col=portrait?i%columns:Math.floor(i/rows),row=portrait?Math.floor(i/columns):i%rows;
        p.x=(portrait?30:offset)+col*272+116;
        p.y=(portrait?offset:30)+row*144+52;});
      offset+=(portrait?rows*144:columns*272)+100;
    }
    const top=Math.min(...nodes.map(n=>graph.node(n.id).y-52))-24;
    const left=Math.min(...nodes.map(n=>graph.node(n.id).x-116))-24;
    for(const edge of edges){const a=graph.node(edge.source),b=graph.node(edge.target);
      const s=portrait?{x:a.x,y:a.y+52}:{x:a.x+116,y:a.y},t=portrait?{x:b.x,y:b.y-52}:{x:b.x-116,y:b.y};
      // Corridors stay outside every packed column/row and above/left of the
      // whole graph, so no link cuts through a sibling card.
      const points=portrait?[s,{x:s.x,y:s.y+18},{x:left,y:s.y+18},{x:left,y:t.y-18},{x:t.x,y:t.y-18},t]
        :[s,{x:s.x+18,y:s.y},{x:s.x+18,y:top},{x:t.x-18,y:top},{x:t.x-18,y:t.y},t];
      edge.sourceHandle='source-'+(portrait?'bottom':'right');edge.targetHandle='target-'+(portrait?'top':'left');
      edge.data={...edge.data,points};
    }
    return nodes.map(n=>({...n,position:{x:graph.node(n.id).x-116,y:graph.node(n.id).y-52}}));
  }
  const top=Math.min(...nodes.map(n=>graph.node(n.id).y-52))-24;
  const left=Math.min(...nodes.map(n=>graph.node(n.id).x-116))-24;
  edges.forEach((edge) => {
    const a=graph.node(edge.source),b=graph.node(edge.target);
    let points,sourceSide,targetSide;
    if(!portrait){
      const direction=Math.sign(b.x-a.x);
      sourceSide=direction>=0?'right':'left';targetSide=direction>0?'left':'right';
      const s={x:a.x+(sourceSide==='right'?116:-116),y:a.y},t={x:b.x+(targetSide==='right'?116:-116),y:b.y};
      if(direction&&Math.abs(b.x-a.x)<=312){const mid=(s.x+t.x)/2;points=[s,{x:mid,y:s.y},{x:mid,y:t.y},t];}
      else if(!direction){sourceSide=targetSide='left';s.x=a.x-116;t.x=b.x-116;const x=Math.min(s.x,t.x)-24;points=[s,{x,y:s.y},{x,y:t.y},t];}
      else points=[s,{x:s.x+direction*24,y:s.y},{x:s.x+direction*24,y:top},{x:t.x-direction*24,y:top},{x:t.x-direction*24,y:t.y},t];
    }else{
      const direction=Math.sign(b.y-a.y);
      sourceSide=direction>=0?'bottom':'top';targetSide=direction>0?'top':'bottom';
      const s={x:a.x,y:a.y+(sourceSide==='bottom'?52:-52)},t={x:b.x,y:b.y+(targetSide==='bottom'?52:-52)};
      if(direction&&Math.abs(b.y-a.y)<=184){const mid=(s.y+t.y)/2;points=[s,{x:s.x,y:mid},{x:t.x,y:mid},t];}
      else if(!direction){sourceSide=targetSide='top';s.y=a.y-52;t.y=b.y-52;const y=Math.min(s.y,t.y)-24;points=[s,{x:s.x,y},{x:t.x,y},t];}
      else points=[s,{x:s.x,y:s.y+direction*24},{x:left,y:s.y+direction*24},{x:left,y:t.y-direction*24},{x:t.x,y:t.y-direction*24},t];
    }
    edge.sourceHandle='source-'+sourceSide;edge.targetHandle='target-'+targetSide;
    edge.data = {...edge.data,points};
  });
  return nodes.map((node) => {
    const p = graph.node(node.id);
    return { ...node, position: { x: p.x - 116, y: p.y - 52 } };
  });
}

export function mount(element, options) {
  let instance,
    update,
    current = { nodes: [], edges: [], language: "fr" };
  const root = createRoot(element);
  function Graph() {
    const hudRef = useRef(null);
    const [model, setModel] = useState(current);
    update = setModel;
    const en = model.language === "en";
    useEffect(() => {
      let previousSize=`${element.clientWidth}:${element.clientHeight}`;
      const stageObserver = new ResizeObserver(() => {
        element.style.setProperty('--graph-stage-width', `${element.clientWidth}px`);
        const size=`${element.clientWidth}:${element.clientHeight}`;
        if(size!==previousSize&&element.clientWidth&&element.clientHeight){
          previousSize=size;requestAnimationFrame(()=>api.fit());
        }
      });
      stageObserver.observe(element);
      const observer = new ResizeObserver(([entry]) =>
        element.style.setProperty(
          "--graph-hud-height",
          `${entry.contentRect.height}px`,
        ),
      );
      if (hudRef.current) observer.observe(hudRef.current);
      return () => { observer.disconnect(); stageObserver.disconnect(); };
    }, []);
    useEffect(() => {
      requestAnimationFrame(() => {
        if (model.focus) api.focus(model.focus);
        else if (model.fit) api.fit();
      });
    }, [model]);
    return (
      <ReactFlow
        nodes={model.nodes}
        edges={model.edges}
        nodeTypes={nodeTypes}
        edgeTypes={edgeTypes}
        onInit={(flow) => {
          instance = flow;
          element._flow = api;
        }}
        onNodesChange={(changes) => {
          current = {
            ...current,
            fit: changes.some(c=>c.type==='position'||c.type==='select')?false:current.fit,
            nodes: applyNodeChanges(changes, current.nodes),
          };
          setModel(current);
        }}
        onNodeClick={(_, node) => options.onSelect(node.data.original)}
        onMove={(_, viewport) => options.onMove?.(viewport)}
        onMoveStart={(event) => {if(event)current.fit=false;}}
        nodesConnectable={false}
        deleteKeyCode={null}
        minZoom={0.001}
        maxZoom={3}
        ariaLabelConfig={
          en
            ? {
                "node.a11yDescription.default":
                  "Press Enter or Space to select a node. Arrow keys move it; Escape clears selection.",
                "edge.a11yDescription.default":
                  "Press Enter or Space to select a relationship; Escape clears selection.",
              }
            : {
                "node.a11yDescription.default":
                  "Entrée ou Espace sélectionne un nœud. Les flèches le déplacent ; Échap annule la sélection.",
                "node.a11yDescription.keyboardDisabled":
                  "Entrée ou Espace sélectionne un nœud ; Échap annule la sélection.",
                "node.a11yDescription.ariaLiveMessage": ({ direction, x, y }) =>
                  `Nœud déplacé vers ${{ left: "la gauche", right: "la droite", up: "le haut", down: "le bas" }[direction] || direction}. Position x : ${x}, y : ${y}`,
                "edge.a11yDescription.default":
                  "Entrée ou Espace sélectionne une relation ; Échap annule la sélection.",
                "controls.ariaLabel": "Commandes du graphe",
                "handle.ariaLabel": "Point de liaison",
                "controls.zoomIn.ariaLabel": "Agrandir",
                "controls.zoomOut.ariaLabel": "Réduire",
                "controls.fitView.ariaLabel": "Ajuster la vue",
                "controls.interactive.ariaLabel": "Activer les interactions",
                "minimap.ariaLabel": "Carte du graphe",
              }
        }
      >
        <MiniMap
          pannable
          zoomable
          nodeColor={(node) =>
            node.data.tone === "focus"
              ? "#145b4b"
              : node.type === "sheet"
                ? "#c6dff5"
                : "#cde9df"
          }
          ariaLabel={en ? "Graph overview" : "Carte du graphe"}
        />
        <Controls orientation={model.portrait?'horizontal':'vertical'} showInteractive={false} fitViewOptions={api.fitOptions()} />
        <Background gap={22} size={1} color="#ccd9d4" />
        <Panel ref={hudRef} position="top-left" className="graph-hud">
          <Dock elements={options.toolbar} />
        </Panel>
        <Panel position="top-left" className="graph-breadcrumb-panel">
          <Dock elements={options.breadcrumbs} />
        </Panel>
        <Panel position="bottom-left" className="graph-status-panel">
          <span className="flow-mobile-status" role="status">{model.nodes.length} {en?'cards':'cartes'} · {model.edges.length} {en?'links':'liens'}</span>
          <Dock elements={options.status} />
        </Panel>
      </ReactFlow>
    );
  }
  const api = {
    render(
      nodes,
      edges,
      { language = "fr", selected, fit = true, focus = null } = {},
    ) {
      const mapped = nodes.map((n) => ({
        id: n.id,
        type: n.type || "cell",
        position: n.position || { x: 0, y: 0 },
        selected: n.nodeId === selected || n.id === selected,
        data: {
          ...n,
          original: n.original || n,
          kindLabel: n.kindLabel || "",
          title: n.title || n.label || n.id,
          subtitle: n.subtitle || "",
          tone: n.tone || "",
          portrait: element.clientWidth < 550,
        },
      }));
      const frequencies=new Map();
      for(const e of edges)frequencies.set(e.weight,(frequencies.get(e.weight)||0)+1);
      const links = edges.map((e, i) => ({
        ...e,
        id: e.id || `edge-${i}`,
        type: "layered",
        markerEnd: { type: MarkerType.ArrowClosed },
        label: e.label ?? (e.weight>1&&frequencies.get(e.weight)<8 ? String(e.weight) : undefined),
        className: e.boundary ? "lineage-boundary" : e.membership ? "membership" : "",
      }));
      const aspect=Math.max(.3,element.clientWidth/Math.max(1,element.clientHeight));
      const padding=api.fitOptions().padding;
      const placed = layered(mapped, links, element.clientWidth<550,aspect,
        element.clientWidth-96,Math.max(60,element.clientHeight-parseFloat(padding.top)-parseFloat(padding.bottom)));
      current = { nodes: placed, edges: links, language, fit, focus,portrait:element.clientWidth<550 };
      flushSync(() => update(current));
    },
    getNodes: () => current.nodes,
    getEdges: () => current.edges,
    getNode: (id) => current.nodes.find((n) => n.id === id),
    getNodePoint: (id) => {
      const n = api.getNode(id),
        v = api.getViewport();
      return n
        ? {
            x: (n.position.x + 116) * v.zoom + v.x,
            y: (n.position.y + 52) * v.zoom + v.y,
          }
        : { x: -1, y: -1 };
    },
    getNodeBounds: (id, rendered = false) => {
      const n = api.getNode(id);
      if (!n) return null;
      const v = rendered ? api.getViewport() : { x: 0, y: 0, zoom: 1 };
      return {
        x1: n.position.x * v.zoom + v.x,
        y1: n.position.y * v.zoom + v.y,
        x2: (n.position.x + 232) * v.zoom + v.x,
        y2: (n.position.y + 104) * v.zoom + v.y,
      };
    },
    getViewport: () => instance?.getViewport() || { x: 0, y: 0, zoom: 1 },
    setViewport: (viewport) => {current.fit=false;return instance?.setViewport(viewport);},
    setPositions: (positions) => {
      current = {
        ...current,
        nodes: current.nodes.map((n) => ({
          ...n,
          position: positions[n.id] || n.position,
        })),
        fit: false,
      };
      flushSync(() => update(current));
    },
    fitOptions: () => {
      const hud =
        element.querySelector(".graph-hud")?.getBoundingClientRect().height ||
        50;
      const breadcrumbs = element.querySelector("#graph-breadcrumbs");
      const panel=element.querySelector('.graph-breadcrumb-panel');
      const trail=breadcrumbs&&!breadcrumbs.hidden&&panel
        ? panel.getBoundingClientRect().bottom-element.getBoundingClientRect().top+12 : 0;
      const overlays=['.graph-status-panel','.react-flow__minimap','.react-flow__controls'].map(s=>element.querySelector(s)?.getBoundingClientRect().height||0);
      const top=Math.max(hud+26,trail),bottom=Math.min(Math.max(...overlays)+16,element.clientHeight*.34);
      return {
        padding: {
          top: `${top}px`,
          right: "48px",
          bottom: `${bottom}px`,
          left: "48px",
        },
        minZoom: 0.001,
        maxZoom: 1.2,
        duration: 0,
      };
    },
    fit: () => instance?.fitView(api.fitOptions()),
    focus: () => instance?.fitView(api.fitOptions()),
    select: (id) => {
      const n = api.getNode(id);
      if (n) options.onSelect(n.data.original);
    },
    zoom: (factor) => {current.fit=false;return instance?.zoomTo((instance?.getZoom() || 1) * factor);},
    panBy: (delta) => {
      current.fit=false;
      const v = api.getViewport();
      instance?.setViewport({ ...v, x: v.x + delta.x, y: v.y + delta.y });
    },
    resize: () => {},
    destroy: () => {
      root.unmount();
      delete element._flow;
    },
  };
  flushSync(() =>
    root.render(
      <ReactFlowProvider>
        <Graph />
      </ReactFlowProvider>,
    ),
  );
  return api;
}

const SEARCH_BOUNDARY_LABELS = {
  input_constant: ["Origine valide · donnée source", "Valid origin · source value"],
  unresolved_external: ["Classeur externe introuvable", "External workbook not found"],
  ambiguous_external: ["Plusieurs classeurs possibles", "Multiple workbooks are possible"],
  unsupported_reference: ["Référence non prise en charge", "Unsupported reference"],
  cache_only: ["Cache seul · source indisponible", "Cache only · source unavailable"],
  macro: ["Arrêt · macro non exécutée", "Stop · macro not executed"],
  power_query: ["Arrêt · Power Query non exécuté", "Stop · Power Query not executed"],
  cycle: ["Arrêt · cycle détecté", "Stop · cycle detected"],
  source_unavailable: ["Source indisponible", "Source unavailable"],
  source_index_incomplete: ["Index source incomplet", "Source index incomplete"],
  possible_unresolved_reader: ["Lecteur potentiel non résolu", "Possible unresolved reader"],
  depth: ["Arrêt · limite de profondeur", "Stop · depth limit"],
  node_budget: ["Arrêt · limite de nœuds", "Stop · node budget"],
  edge_budget: ["Arrêt · limite de relations", "Stop · edge budget"],
  workbook_budget: ["Arrêt · limite de classeurs", "Stop · workbook budget"],
  time_budget: ["Arrêt · limite de durée", "Stop · time budget"],
  response_bytes: ["Arrêt · limite de réponse", "Stop · response size limit"],
  link_budget: ["Arrêt · limite de liens", "Stop · link budget"],
  link_build_time: ["Arrêt · durée de construction des liens", "Stop · link build time"],
  index_rebuild_required: ["Réindexation requise", "Index rebuild required"],
};

function searchText(value, english) {
  if (Array.isArray(value)) return value[english ? 1 : 0];
  if (value && typeof value === "object") return JSON.stringify(value);
  return value == null ? "" : String(value);
}

function LibrarySearch({ apiBase = "api/", language = "fr" }) {
  const english = language === "en";
  const text = (fr, en) => (english ? en : fr);
  const [query, setQuery] = useState("");
  const [mode, setMode] = useState("library");
  const [projects, setProjects] = useState([]);
  const [projectId, setProjectId] = useState("");
  const [revisionId, setRevisionId] = useState("");
  const [kinds, setKinds] = useState(["cell", "table", "column", "name", "sheet"]);
  const [search, setSearch] = useState(null);
  const [selected, setSelected] = useState(null);
  const [lineage, setLineage] = useState(null);
  const [direction, setDirection] = useState("upstream");
  const [status, setStatus] = useState(null);
  const [links, setLinks] = useState([]);
  const [linkStatus, setLinkStatus] = useState("");
  const [pinTargets, setPinTargets] = useState({});
  const [traceEvents, setTraceEvents] = useState([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [attachment, setAttachment] = useState(null);
  const [uploadProgress, setUploadProgress] = useState("");
  const [occurrences, setOccurrences] = useState({});
  const canvas = useRef(null);
  const flow = useRef(null);

  const request = async (path, options = {}) => {
    const response = await fetch(`${apiBase}${path}`, {
      ...options,
      headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    });
    let data;
    try {
      data = await response.json();
    } catch {
      throw new Error(text("Réponse du serveur illisible.", "Unreadable server response."));
    }
    if (!response.ok) throw new Error(data.error?.message || data.message || `HTTP ${response.status}`);
    return data;
  };
  const scopedParams = (scope = true) => {
    const params = new URLSearchParams();
    if (scope && attachment?.scopeId) params.set("scope", attachment.scopeId);
    return params;
  };
  const refreshStatus = async () => {
    const params = scopedParams();
    const value = await request(`library/status${params.size ? `?${params}` : ""}`);
    setStatus(value);
    return value;
  };
  const refreshLinks = async () => {
    const params = scopedParams();
    params.set("limit", "100");
    if (linkStatus) params.set("status", linkStatus);
    const value = await request(`library/links?${params}`);
    setLinks(value.links || []);
  };
  useEffect(() => {
    let active = true;
    Promise.all([request("projects"), refreshStatus()])
      .then(([projectData]) => {
        if (active) setProjects(projectData.projects || []);
      })
      .catch((cause) => active && setError(cause.message));
    refreshLinks().catch((cause) => active && setError(cause.message));
    return () => { active = false; };
  }, [attachment?.scopeId, linkStatus]);

  const addTrace = (type, trace, detail = "") => {
    setTraceEvents((events) => [...events, { type, trace, detail, at: new Date().toISOString() }]);
  };
  const apiScope = () => {
    const params = scopedParams();
    if (mode === "project" && projectId) params.set("project", projectId);
    if (mode === "attachment" && attachment?.project?.id) params.set("project", attachment.project.id);
    if (mode === "revision" && revisionId) {
      params.set("revision", revisionId);
      if (projectId) params.set("project", projectId);
    }
    return params;
  };
  const performSearch = async (offset = 0, generation = "") => {
    setBusy(true);
    setError("");
    setLineage(null);
    setSelected(null);
    try {
      const params = apiScope();
      params.set("q", query);
      params.set("kinds", kinds.join(","));
      params.set("limit", "50");
      params.set("offset", String(offset));
      if (generation) params.set("generation", generation);
      const result = await request(`library/search?${params}`);
      setSearch(result);
      setOccurrences({});
      addTrace("search", result.trace, `${result.results?.length || 0}`);
    } catch (cause) {
      setError(cause.message);
    } finally {
      setBusy(false);
    }
  };
  const selectResult = async (hit) => {
    setSelected(hit);
    setLineage(null);
    if (!hit.handle) return;
    setBusy(true);
    setError("");
    try {
      const params = scopedParams();
      params.set("handle", hit.handle);
      params.set("direction", direction);
      params.set("depth", "20");
      params.set("limit", "200");
      const result = await request(`library/lineage?${params}`);
      setLineage(result);
      addTrace("lineage", result.trace, direction);
    } catch (cause) {
      setError(cause.message);
    } finally {
      setBusy(false);
    }
  };
  useEffect(() => {
    if (!canvas.current || !lineage) return undefined;
    const instance = mount(canvas.current, {
      onSelect: (node) => {
        const found = (lineage.nodes || []).find((item) => item.id === node.id);
        if (found) setSelected(found);
      },
    });
    flow.current = instance;
    const nodes = (lineage.nodes || []).map((node, index) => {
      const rev = node.revision || node.workbookRevision || {};
      const file = node.filename || node.projectName || rev.projectId || node.projectId || text("Classeur", "Workbook");
      const hash = rev.sha256 || node.sha256 || "";
      const place = [node.sheet, node.address || node.cell].filter(Boolean).join(" · ");
      return {
        id: node.id,
        nodeId: node.id,
        kind: node.kind,
        kindLabel: `${file} · ${hash.slice(0, 8)} · ${node.kind || ""}`,
        title: node.label || node.address || node.id,
        subtitle: place,
        valueLabel: node.value === null || node.value === undefined ? "" : searchText(node.value, english),
        tone: index === 0 ? "focus" : "",
        original: node,
      };
    });
    const visibleBoundaries = (lineage.boundaries || []).map((boundary, index) => {
      const reason = boundary.reason || boundary.type || "source_unavailable";
      const label = SEARCH_BOUNDARY_LABELS[reason] || [
        `Arrêt · ${reason.replaceAll("_", " ")}`,
        `Stop · ${reason.replaceAll("_", " ")}`,
      ];
      return {
        id: `boundary:${index}`,
        kind: "boundary",
        kindLabel: text("Limite de la piste", "Lineage boundary"),
        title: searchText(label, english),
        subtitle: [boundary.sheet, boundary.address, boundary.detail].filter(Boolean).join(" · "),
        tone: reason === "input_constant" ? "outside" : "boundary",
        original: boundary,
      };
    });
    const graphNodes = [...nodes, ...visibleBoundaries];
    const nodeIds = new Set(nodes.map((node) => node.id));
    const edges = (lineage.edges || []).filter((edge) => nodeIds.has(edge.source) && nodeIds.has(edge.target)).map((edge) => ({
      ...edge,
      label: edge.kind === "external_link" ? text("lien externe", "external link") : edge.kind,
    }));
    visibleBoundaries.forEach((boundary, index) => {
      const fact = (lineage.boundaries || [])[index];
      const reached = fact.nodeId || fact.node || fact.sourceNode || fact.at;
      if (!nodeIds.has(reached)) return;
      edges.push({
        id: `boundary-edge:${index}`,
        source: boundary.tone === "outside" ? reached : boundary.id,
        target: boundary.tone === "outside" ? boundary.id : reached,
        boundary: true,
        label: text("limite", "boundary"),
      });
    });
    const renderFrame = requestAnimationFrame(() => {
      if (flow.current === instance) {
        instance.render(graphNodes, edges, { language: english ? "en" : "fr" });
      }
    });
    return () => {
      cancelAnimationFrame(renderFrame);
      instance.destroy();
      flow.current = null;
    };
  }, [lineage, english]);

  const toggleKind = (kind) => setKinds((selectedKinds) =>
    selectedKinds.includes(kind)
      ? selectedKinds.filter((item) => item !== kind)
      : [...selectedKinds, kind],
  );
  const addOccurrences = async (hit) => {
    const group = hit.occurrenceGroup;
    if (!group) return;
    const currentPage = occurrences[hit.handle] || { rows: hit.occurrences || [], after: 0 };
    const params = scopedParams();
    params.set("generation", group.generation);
    params.set("group", group.id);
    params.set("after", String(currentPage.after));
    params.set("limit", "50");
    try {
      const page = await request(`library/occurrences?${params}`);
      setOccurrences((state) => ({
        ...state,
        [hit.handle]: {
          rows: [...currentPage.rows, ...(page.results || [])],
          after: page.pagination?.nextAfter ?? currentPage.after,
          done: page.pagination?.nextAfter == null,
        },
      }));
      addTrace("occurrences", { generation: group.generation, capsHit: page.coverage?.capsHit || [], complete: page.coverage?.complete }, `${page.results?.length || 0}`);
    } catch (cause) {
      setError(cause.message);
    }
  };
  const attachFile = async (file) => {
    setBusy(true);
    setError("");
    let uploadId = "";
    try {
      const started = await request("uploads", {
        method: "POST",
        body: JSON.stringify({
          workbook: { name: file.name, size: file.size },
          name: file.name,
          searchScoped: true,
          budget: { seconds: 600, memoryMb: 512 },
        }),
      });
      uploadId = started.upload.id;
      const chunkBytes = started.upload.chunkBytes || 8 * 1024 * 1024;
      for (let offset = 0; offset < file.size; offset += chunkBytes) {
        setUploadProgress(`${text("Envoi", "Uploading")} ${Math.min(file.size, offset + chunkBytes)} / ${file.size}`);
        const response = await fetch(`${apiBase}uploads/${uploadId}/files/0?offset=${offset}`, {
          method: "PUT",
          headers: { "Content-Type": "application/octet-stream", "X-Upload-Offset": String(offset) },
          body: file.slice(offset, Math.min(file.size, offset + chunkBytes)),
        });
        if (!response.ok) throw new Error(`${text("Envoi interrompu", "Upload interrupted")} (${response.status})`);
      }
      setUploadProgress(text("Import structurel en cours…", "Structural import running…"));
      const imported = await request(`uploads/${uploadId}/complete`, { method: "POST", body: "{}" });
      let task = imported.task;
      while (task.status !== "succeeded" && task.status !== "failed" && task.status !== "cancelled") {
        setUploadProgress(`${text("Import", "Import")} · ${task.progress?.phase || task.status}`);
        await new Promise((resolve) => setTimeout(resolve, 350));
        task = (await request(`tasks/${task.id}`)).task;
      }
      if (task.status !== "succeeded") throw new Error(task.error?.message || text("Import échoué.", "Import failed."));
      const scopeId = imported.project.searchScopeId || uploadId;
      setAttachment({ project: imported.project, scopeId });
      setProjectId(imported.project.id);
      setMode("attachment");
      setUploadProgress(text("Fichier joint à cette recherche uniquement.", "File attached to this search only."));
      addTrace("upload", null, text("Import terminé · périmètre temporaire", "Import completed · temporary scope"));
      await refreshStatus();
      await refreshLinks();
    } catch (cause) {
      if (uploadId) {
        try { await request(`uploads/${uploadId}`, { method: "DELETE", body: "{}" }); } catch { /* Report the original upload/import error. */ }
      }
      setError(cause.message);
    } finally {
      setBusy(false);
    }
  };
  const keepAttachment = async () => {
    if (!attachment) return;
    try {
      await request(`library/attachments/${attachment.scopeId}/keep`, { method: "POST", body: "{}" });
      setAttachment(null);
      setMode("project");
      setUploadProgress(text("Classeur conservé dans la bibliothèque.", "Workbook kept in the library."));
      const data = await request("projects");
      setProjects(data.projects || []);
      await refreshStatus();
      await refreshLinks();
    } catch (cause) {
      setError(cause.message);
    }
  };
  const removeAttachment = async () => {
    if (!attachment) return;
    try {
      await request(`library/attachments/${attachment.scopeId}`, { method: "DELETE", body: "{}" });
      setAttachment(null);
      setMode("library");
      setSearch(null);
      setSelected(null);
      setLineage(null);
      setUploadProgress("");
    } catch (cause) {
      setError(cause.message);
    }
  };
  const pin = async (link, remove = false) => {
    try {
      const params = scopedParams();
      if (remove) {
        await request(`library/pins/${encodeURIComponent(link.id)}${params.size ? `?${params}` : ""}`, { method: "DELETE", body: "{}" });
      } else {
        const [targetProject, targetSha256] = (pinTargets[link.id] || "").split(":");
        if (!targetProject || !targetSha256) throw new Error(text("Choisissez une révision cible.", "Choose a target revision."));
        await request("library/pins", {
          method: "POST",
          body: JSON.stringify({ linkId: link.id, targetProject, targetSha256, ...(attachment ? { scope: attachment.scopeId } : {}) }),
        });
      }
      await refreshLinks();
      await refreshStatus();
    } catch (cause) {
      setError(cause.message);
    }
  };
  const results = search?.results || [];
  const revisions = status?.revisions || [];
  const complete = search?.coverage?.complete;
  const caps = search?.coverage?.capsHit || search?.trace?.capsHit || [];
  const lineageCoverage = lineage?.coverage;
  const lineageNodes = lineage?.nodes || [];
  const lineageEdges = lineage?.edges || [];
  const lineageBoundaries = lineage?.boundaries || [];
  const lineageNodeLabels = new Map(lineageNodes.map((node) => [
    node.id,
    [
      node.filename || node.revision?.projectId || node.projectId,
      node.revision?.sha256?.slice(0, 8),
      node.sheet && `${node.sheet}!${node.address || node.cell || ""}`,
      node.kind,
    ].filter(Boolean).join(" · "),
  ]));
  const traceTime = (value) => value == null ? "" : `${value} ms`;
  const revisionLabel = (revision) => {
    const sha = revision?.sha256 || "";
    return `${revision?.filename || revision?.projectId || ""} · ${sha.slice(0, 8)}`;
  };
  return (
    <div className="library-search-view">
      <header className="library-search-heading">
        <div><p className="section-label">{text("Bibliothèque de classeurs", "Workbook library")}</p><h2>{text("Rechercher", "Search")}</h2></div>
        <p>{text("La valeur et la piste proviennent des index chargés. Leur pertinence reste à confirmer.", "Values and trails come from loaded indexes. Relevance still needs confirmation.")}</p>
      </header>
      <section className="library-catalog-state" aria-live="polite">
        <strong>{text("Catalogue", "Catalog")}</strong>
        <span>{status?.buildState?.status || status?.status || text("État inconnu", "Unknown state")}</span>
        {status?.stale && <span>{text("Périmé ou en attente de mise à jour", "Stale or waiting for refresh")}</span>}
        {status?.buildState?.errorType && <span>{text("Échec", "Failed")}: {status.buildState.errorType}</span>}
      </section>
      <section className="library-search-controls" aria-label={text("Recherche dans la bibliothèque", "Search the library")}>
        <form onSubmit={(event) => { event.preventDefault(); performSearch(); }}>
          <label htmlFor="library-search-query">{text("Requête", "Query")}</label>
          <div className="library-search-form-row">
            <input id="library-search-query" type="search" value={query} onChange={(event) => setQuery(event.target.value)} />
            <button type="submit" className="primary" disabled={busy}>{text("Rechercher", "Search")}</button>
          </div>
        </form>
        <fieldset className="library-scope">
          <legend>{text("Portée", "Scope")}</legend>
          {[
            ["library", text("Toute la bibliothèque", "Whole library")],
            ["project", text("Classeurs", "Workbooks")],
            ["revision", text("Révision", "Revision")],
            ["attachment", text("Fichier joint", "Attached file")],
          ].map(([value, label]) => (
            <button key={value} type="button" className={mode === value ? "is-active" : ""} aria-pressed={mode === value} disabled={value === "attachment" && !attachment} onClick={() => setMode(value)}>{label}</button>
          ))}
          {mode === "project" && <label>{text("Classeur", "Workbook")}<select value={projectId} onChange={(event) => setProjectId(event.target.value)}><option value="">{text("Choisir…", "Choose…")}</option>{projects.map((project) => <option key={project.id} value={project.id}>{project.filename} · {project.id.slice(0, 8)}</option>)}</select></label>}
          {mode === "revision" && <label>{text("Révision", "Revision")}<select value={revisionId} onChange={(event) => setRevisionId(event.target.value)}><option value="">{text("Choisir…", "Choose…")}</option>{revisions.map((revision) => <option key={revision.revision?.id || revision.projectId} value={revision.revision?.id}>{revisionLabel({ ...revision, sha256: revision.revision?.sha256 })} · {revision.projectId.slice(0, 8)}</option>)}</select></label>}
          {mode === "attachment" && attachment && <span>{attachment.project.filename} · {attachment.project.id.slice(0, 8)}</span>}
        </fieldset>
        <fieldset className="library-kind-filter">
          <legend>{text("Types", "Kinds")}</legend>
          {["cell", "sheet", "name", "table", "column"].map((kind) => <label key={kind}><input type="checkbox" checked={kinds.includes(kind)} onChange={() => toggleKind(kind)} />{kind}</label>)}
        </fieldset>
        <div className="library-attach-control">
          <label htmlFor="library-attach-file">{text("Joindre un classeur à cette recherche", "Attach a workbook to this search")}</label>
          <input id="library-attach-file" type="file" accept=".xlsx,.xlsm,.xlam,.xla" disabled={busy} onChange={(event) => { const file = event.target.files?.[0]; if (file) attachFile(file); event.target.value = ""; }} />
          {uploadProgress && <p role="status">{uploadProgress}</p>}
          {attachment && <div className="library-attachment-actions"><button type="button" onClick={keepAttachment}>{text("Conserver dans la bibliothèque", "Keep in library")}</button><button type="button" onClick={removeAttachment}>{text("Retirer le fichier joint", "Remove attached file")}</button></div>}
        </div>
      </section>
      {error && <p className="notice error" role="alert">{error}</p>}
      <div className="library-search-columns">
        <section className="library-results" aria-labelledby="library-results-heading">
          <h3 id="library-results-heading">{text("Résultats", "Results")}</h3>
          {search && results.length === 0 && <p className="library-empty">{text("Aucun résultat affiché pour cette requête et cette portée.", "No results are displayed for this query and scope.")}</p>}
          {results.map((hit) => {
            const labels = hit.labels || [];
            const occurrenceState = occurrences[hit.handle];
            const shownOccurrences = occurrenceState?.rows || hit.occurrences || [];
            return <article className={`library-result ${selected?.handle === hit.handle ? "is-selected" : ""}`} key={`${hit.revision?.id}:${hit.kind}:${hit.address}:${hit.sheet}:${hit.entryId || ""}`}>
              <button type="button" className="library-result-select" onClick={() => selectResult(hit)} aria-label={`${text("Afficher la piste pour", "Show lineage for")} ${hit.filename || ""} ${hit.sheet || ""}!${hit.address || ""}`}>
                <strong className="library-result-value">{searchText(hit.value, english)}</strong>
                <span>{hit.filename} · {hit.revision?.sha256?.slice(0, 8)} · {hit.revision?.projectId?.slice(0, 8)}</span>
                <span>{hit.sheet}{hit.address ? `!${hit.address}` : ""} · {hit.kind}</span>
              </button>
              {!!labels.length && <ul className="library-labels">{labels.map((label, index) => {
                const declared = ["table_column", "defined_name"].includes(label.method) || ["table_column", "defined_name"].includes(label.role);
                const heuristic = label.certainty === "heuristic" || !declared;
                return <li key={`${label.text}:${label.role}:${index}`}><span>{label.role || label.method}</span><strong>{label.text}</strong><small>{declared ? text("Déclaré", "Declared") : heuristic ? text("Inféré · heuristique", "Inferred · heuristic") : label.method}</small></li>;
              })}</ul>}
              {hit.occurrenceCount !== undefined && <div className="library-occurrences">
                <span>{hit.occurrenceCount} {text("occurrences", "occurrences")} · {hit.occurrenceCountComplete ? text("comptage complet", "complete count") : text("comptage partiel", "partial count")}</span>
                {shownOccurrences.map((occurrence) => <button key={occurrence.handle || occurrence.address} type="button" onClick={() => selectResult({ ...hit, handle: occurrence.handle, sheet: occurrence.sheet, address: occurrence.address })}>{occurrence.sheet}!{occurrence.address}</button>)}
                {hit.occurrenceGroup && !occurrenceState?.done && <button type="button" onClick={() => addOccurrences(hit)}>{text("Afficher la page suivante", "Show next page")}</button>}
              </div>}
            </article>;
          })}
          {search?.pagination?.hasMore && <button type="button" onClick={() => performSearch((search.pagination.offset || 0) + search.pagination.limit, search.generation)}>{text("Résultats suivants", "Next results")}</button>}
        </section>
        <section className="library-lineage" aria-labelledby="library-lineage-heading">
          <div className="library-section-heading"><h3 id="library-lineage-heading">{text("Piste de lignage", "Lineage trail")}</h3>
            <label>{text("Sens", "Direction")}<select value={direction} onChange={(event) => { setDirection(event.target.value); if (selected) selectResult(selected); }}><option value="upstream">{text("En amont", "Upstream")}</option><option value="downstream">{text("En aval", "Downstream")}</option></select></label>
          </div>
          {selected && lineageCoverage && <p className={`library-verdict ${lineageCoverage.complete ? "is-complete" : "is-incomplete"}`} role="status">
            {lineageCoverage.exhaustive ? text("Parcours exhaustif dans le périmètre chargé.", "Traversal exhaustive within the loaded scope.") : text("Parcours non exhaustif.", "Traversal is not exhaustive.")}
            {" "}{lineageCoverage.complete ? text("Lignage complet.", "Lineage complete.") : text("Lignage incomplet.", "Lineage incomplete.")}
            {lineageCoverage.boundaries?.length > 0 && ` ${text("Arrêts", "Stops")}: ${lineageCoverage.boundaries.map((reason) => searchText(SEARCH_BOUNDARY_LABELS[reason] || reason, english)).join(", ")}.`}
            {lineageCoverage.capsHit?.length > 0 && ` ${text("Limites", "Caps")}: ${lineageCoverage.capsHit.join(", ")}.`}
          </p>}
          {lineage && <div className="library-diagnostics">
            {(lineage.edges || []).flatMap((edge, index) => [
              ...(edge.resolutionStatus === "resolved_by_name" ? [<p key={`resolution:${index}`}>{text("Résolu par nom : la version peut différer de celle visée lors de l’enregistrement du lien.", "Resolved by name: the version may differ from the one the link was saved against.")}</p>] : []),
              ...(edge.cacheDiagnostic || []).map((diagnostic, diagnosticIndex) => <p key={`cache:${index}:${diagnosticIndex}`}>{text("Diagnostic cache", "Cache diagnostic")} · {diagnostic.address}: {searchText(diagnostic.cached, english)} → {searchText(diagnostic.current, english)} · {diagnostic.agreement === false ? text("différent", "different") : diagnostic.agreement === true ? text("identique", "same") : text("comparaison inconnue", "comparison unknown")}</p>),
            ])}
          </div>}
          {lineage ? <div className="library-graph-layout">
            <div className="library-graph-canvas" ref={canvas} role="application" aria-label={text("Graphe de lignage. Une alternative en liste suit.", "Lineage graph. A list alternative follows.")} />
            {lineageBoundaries.length > 0 && <p className="library-graph-note">{text("Le graphe ajoute une carte d’origine ou d’arrêt et une relation en pointillés pour chaque frontière ; ces éléments synthétiques ne sont pas inclus dans les comptes API de la trace.", "The graph adds a synthetic origin or stop card and a dashed relationship for each boundary; these are not included in the trace’s API counts.")}</p>}
            <details className="library-accessible-graph" open>
              <summary>{text("Alternative accessible · nœuds et relations", "Accessible alternative · nodes and relationships")}</summary>
              <ul>
                {lineageNodes.map((node) => <li key={node.id}>{lineageNodeLabels.get(node.id)}</li>)}
                {lineageEdges.map((edge, index) => <li key={`${edge.source}:${edge.target}:${index}`}>{lineageNodeLabels.get(edge.source) || edge.source} → {lineageNodeLabels.get(edge.target) || edge.target} · {text("relation orientée", "directed relationship")} · {edge.kind}{edge.resolutionStatus ? ` · ${edge.resolutionStatus}` : ""}</li>)}
                {lineageBoundaries.map((boundary, index) => {
                  const nodeId = boundary.nodeId || boundary.node || boundary.sourceNode;
                  return <li key={`boundary:${index}`}>{text("Arrêt", "Stop")} · {searchText(SEARCH_BOUNDARY_LABELS[boundary.reason] || boundary.reason, english)} · {lineageNodeLabels.get(nodeId) || nodeId || ""}</li>;
                })}
              </ul>
            </details>
          </div> : <p className="library-empty">{text("Choisissez un résultat pour suivre ses sources.", "Choose a result to follow its sources.")}</p>}
        </section>
      </div>
      <details className="library-links" open>
        <summary>{text("Liens externes", "External links")}</summary>
        <label>{text("État du lien", "Link status")}<select value={linkStatus} onChange={(event) => setLinkStatus(event.target.value)}><option value="">{text("Tous", "All")}</option>{["ambiguous", "unresolved", "pinned", "resolved_by_name"].map((value) => <option key={value}>{value}</option>)}</select></label>
        {links.length === 0 ? <p>{text("Aucun lien dans cette portée.", "No links in this scope.")}</p> : links.map((link) => {
          const pinned = link.status === "pinned" || link.pin;
          const targetOptions = revisions.filter((item) => item.revision?.sha256);
          return <article key={link.id} className="library-link">
            <div><strong>{link.filename || link.sourceFilename || link.sourceProject || text("Classeur source", "Source workbook")}</strong><span>{link.path || link.targetFilename || link.target || link.status}</span><span>{link.status}{link.stalePin ? ` · ${text("épingle périmée", "stale pin")}` : ""}</span></div>
            <label>{text("Révision cible", "Target revision")}<select aria-label={`${text("Révision cible pour", "Target revision for")} ${link.id}`} value={pinTargets[link.id] || ""} onChange={(event) => setPinTargets((state) => ({ ...state, [link.id]: event.target.value }))}><option value="">{text("Choisir une révision", "Choose a revision")}</option>{targetOptions.map((item) => <option key={item.revision.id} value={`${item.projectId}:${item.revision.sha256}`}>{revisionLabel({ filename: item.filename, projectId: item.projectId, sha256: item.revision.sha256 })}</option>)}</select></label>
            <button type="button" onClick={() => pin(link)}>{pinned ? text("Épingler la révision choisie", "Pin chosen revision") : text("Épingler", "Pin")}</button>
            {pinned && <button type="button" onClick={() => pin(link, true)}>{text("Retirer l’épingle", "Unpin")}</button>}
          </article>;
        })}
      </details>
      <section className="library-trace" aria-labelledby="library-trace-heading">
        <h3 id="library-trace-heading">{text("Trace de recherche et de lignage", "Search and lineage trace")}</h3>
        {traceEvents.length === 0 && <p>{text("Aucune requête effectuée. La trace apparaîtra ici, y compris si aucun résultat n’est trouvé.", "No request yet. The trace will appear here, including when a search has no results.")}</p>}
        <ol>{traceEvents.map((event, index) => {
          const trace = event.trace || {};
          const scope = trace.scope || {};
          const counts = trace.counts || {};
          return <li key={`${event.type}:${index}`}>
            <strong>{event.type === "search" ? text("Recherche", "Search") : event.type === "lineage" ? text("Lignage", "Lineage") : event.type === "occurrences" ? text("Occurrences", "Occurrences") : text("Import", "Import")}</strong>
            <span>{event.type === "search" ? `${text("Termes normalisés", "Normalized terms")}: ${(trace.terms || []).join(", ") || text("aucun", "none")} · ${scope.project || text("toute la bibliothèque", "whole library")} ${scope.revision || ""} · ${(scope.kinds || []).join(", ")}` : event.type === "lineage" ? `${event.detail} · ${text("révisions", "revisions")}: ${(trace.revisions || []).map(revisionLabel).join(", ")}` : event.detail}</span>
            <small>{trace.generation ? `gen ${trace.generation}` : ""}{trace.backend ? ` · ${trace.backend}` : ""}{trace.tool ? ` · ${trace.tool}` : ""}{trace.durationMs !== undefined ? ` · ${traceTime(trace.durationMs)}` : ""}{Object.keys(counts).length ? ` · ${Object.entries(counts).map(([key, value]) => `${key}: ${value}`).join(" · ")}` : ""}{trace.capsHit?.length ? ` · ${text("limites", "caps")}: ${trace.capsHit.join(", ")}` : ""}{event.type === "search" ? ` · ${trace.complete === false ? text("recherche tronquée", "search truncated") : text("recherche complète", "search complete")}` : ""}</small>
          </li>;
        })}</ol>
        <p className={`library-trace-last ${complete === false ? "is-incomplete" : ""}`}>{search ? (complete === false ? text("La recherche dans ce périmètre est incomplète ou tronquée.", "The scoped search is incomplete or truncated.") : text("La recherche du contenu indexé est complète dans ce périmètre.", "The indexed-content search is complete in this scope.")) : text("Aucune recherche exécutée.", "No search performed.")} {text("Une recherche complète du contenu indexé ne prouve pas l’absence d’une valeur dans les classeurs.", "A complete search of indexed content does not prove that a value is absent from workbooks.")}</p>
      </section>
    </div>
  );
}

export function mountLibrarySearch(element, options = {}) {
  const root = createRoot(element);
  let update;
  function SearchRoot() {
    const [language, setLanguage] = useState(options.language || "fr");
    update = setLanguage;
    return <LibrarySearch apiBase={options.apiBase} language={language} />;
  }
  flushSync(() => root.render(<SearchRoot />));
  return {
    setLanguage: (language) => update(language),
    destroy: () => root.unmount(),
  };
}
