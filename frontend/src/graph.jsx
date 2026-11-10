import React, { memo, useEffect, useState, useRef } from "react";
import { AskQuestion, AskTrace } from "./ask.jsx";
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
import { revisionLineageView } from './aggregation.js';
export { compressNeighborhood, revisionLineageView } from './aggregation.js';

const Card = memo(({ data, selected }) => (
  <article
    data-no-i18n=""
    className={`flow-card ${data.tone || ""} ${selected ? "is-selected" : ""}`}
    title={[data.kindLabel, data.title, data.subtitle, data.valueLabel].filter(Boolean).join('\n')}
  >
    {Object.entries({left:Position.Left,right:Position.Right,top:Position.Top,bottom:Position.Bottom}).flatMap(([side,position])=>['source','target'].map(type=><Handle key={`${type}-${side}`} id={`${type}-${side}`} type={type} position={position} isConnectable={false} />))}
    <span className="flow-kind">{data.kindLabel}</span>
    <div className="flow-title-row"><strong>{data.title}</strong>{data.count!==undefined&&<span className="flow-count">{data.count}</span>}</div>
    <span className="flow-subtitle">{data.subtitle}</span>
    {data.valueLabel&&<span className="flow-value">{data.valueLabel}</span>}
  </article>
));
const nodeTypes = { cell: Card, formulaGroup: Card, sheet: Card, cellGroup: Card };
function countLabel(count, english, french, en) {
  const forms = english ? en : french;
  const singular = Number(count) === 1 || (!english && Number(count) === 0);
  return `${count} ${forms[singular ? 0 : 1]}`;
}
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
    center = data?.labelPoint || (data?.bundle ? {x:(points[0].x+points.at(-1).x)/2,y:(points[0].y+points.at(-1).y)/2} : points[Math.floor(points.length / 2)]);
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

function layered(nodes, edges, portrait, aspect, availableWidth, availableHeight, stack = false) {
  if (stack) {
    const positions = new Map(nodes.map((node, index) => [node.id, { x: 30, y: 30 + index * 144 }]));
    for (const [index, edge] of edges.entries()) {
      const a = positions.get(edge.source), b = positions.get(edge.target);
      const lane = 12 - (index % 3) * 5;
      edge.sourceHandle = 'source-left';
      edge.targetHandle = 'target-left';
      // The narrow left corridor holds the path; keep its label in the gap
      // above the lower card so the viewport never clips the relationship.
      edge.data = { ...edge.data, labelPoint: { x: lane + 42, y: Math.max(a.y, b.y) - 20 }, points: [{ x: a.x, y: a.y + 52 }, { x: lane, y: a.y + 52 }, { x: lane, y: b.y + 52 }, { x: b.x, y: b.y + 52 }] };
    }
    return nodes.map(node => ({ ...node, position: positions.get(node.id) }));
  }
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
    const [minimapSize, setMinimapSize] = useState(() =>
      window.matchMedia("(max-width: 650px)").matches
        ? { width: 100, height: 74 }
        : { width: 200, height: 150 },
    );
    update = setModel;
    const en = model.language === "en";
    useEffect(() => {
      const mobile = window.matchMedia("(max-width: 650px)");
      const updateSize = () =>
        setMinimapSize(
          mobile.matches
            ? { width: 100, height: 74 }
            : { width: 200, height: 150 },
        );
      mobile.addEventListener("change", updateSize);
      return () => mobile.removeEventListener("change", updateSize);
    }, []);
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
        className={model.portrait ? "flow-portrait" : undefined}
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
        onNodeClick={(_, node) => {
          if (node.data.original) options.onSelect(node.data.original);
        }}
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
          style={{ width: minimapSize.width, height: minimapSize.height }}
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
          <span className="flow-mobile-status" role="status">{options.relationshipStatus
            ? options.relationshipStatus(model)
            : `${countLabel(model.nodes.length, en, ["carte", "cartes"], ["card", "cards"])} · ${countLabel(model.edges.filter(e => !e.hidden).length, en, ["liaison", "liaisons"], ["link", "links"])}`}</span>
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
        selected: selected != null && (n.nodeId === selected || n.id === selected),
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
        element.clientWidth-96,Math.max(60,element.clientHeight-parseFloat(padding.top)-parseFloat(padding.bottom)), options.stack && element.clientWidth < 550);
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
      const top=options.compact?12:Math.max(hud+26,trail),bottom=Math.min(Math.max(...overlays)+16,element.clientHeight*.34);
      return {
        padding: {
          top: `${top}px`,
          right: options.compact ? "16px" : "48px",
          bottom: `${bottom}px`,
          left: options.compact ? "16px" : "48px",
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
const SEARCH_KIND_LABELS = {
  cell: ["cellule", "cell"],
  sheet: ["feuille", "sheet"],
  name: ["nom défini", "defined name"],
  table: ["tableau", "table"],
  column: ["colonne", "column"],
  formula: ["formule", "formula"],
  input: ["valeur source", "source value"],
  range: ["plage", "range"],
  boundary: ["frontière", "boundary"],
};
const SEARCH_ROLE_LABELS = {
  table_column: ["colonne déclarée", "declared column"],
  defined_name: ["nom défini", "defined name"],
  column: ["colonne", "column"],
  row_context: ["contexte de ligne", "row context"],
  row: ["étiquette de ligne", "row label"],
};
const SEARCH_STATUS_LABELS = {
  ambiguous: ["plusieurs cibles possibles", "multiple possible targets"],
  unresolved: ["cible introuvable", "target not found"],
  pinned: ["cible choisie", "chosen target"],
  resolved_by_name: ["cible trouvée par nom", "target matched by name"],
};
const SEARCH_CAP_LABELS = {
  hits: ["limite de résultats atteinte", "result limit reached"],
  header_candidate_rows: ["analyse des en-têtes limitée", "header scan limited"],
  header_scan_cells: ["analyse des cellules d’en-tête limitée", "header-cell scan limited"],
  header_columns: ["analyse des colonnes d’en-tête limitée", "header-column scan limited"],
  header_row_labels: ["analyse des étiquettes de ligne limitée", "header row-label scan limited"],
  header_bands: ["détection des zones d’en-tête limitée", "header-band detection limited"],
  header_regions: ["zones d’en-tête incomplètes", "header regions incomplete"],
  header_merge_metadata_unavailable: ["métadonnées des cellules fusionnées indisponibles", "merged-cell metadata unavailable"],
  catalog_refresh_pending: ["index en cours de mise à jour", "index is updating"],
  catalog_build_failed: ["la construction du catalogue a échoué", "catalog build failed"],
  catalog_source_snapshot_failed: ["capture des classeurs source échouée", "source-workbook snapshot failed"],
  entries_per_workbook: ["limite d’entrées du classeur", "workbook entry limit"],
  entries_per_sheet: ["limite d’entrées de la feuille", "worksheet entry limit"],
  scan_entries: ["limite de balayage des entrées", "catalog scan limit"],
  numeric_cells_per_workbook: ["limite de cellules numériques du classeur", "workbook numeric-cell limit"],
  text_chars: ["texte indexé tronqué", "indexed text truncated"],
  build_seconds: ["durée de construction du catalogue atteinte", "catalog build time limit reached"],
  catalog_bytes: ["capacité de stockage du catalogue atteinte", "catalog storage limit reached"],
  stale_generation: ["catalogue modifié pendant la recherche", "catalog changed during search"],
  index_rebuild_required: ["réindexation requise", "index rebuild required"],
  page_limit: ["limite de page d’occurrences atteinte", "occurrence page limit reached"],
  query_ms: ["durée maximale de recherche atteinte", "search time limit reached"],
  response_bytes: ["taille maximale de réponse atteinte", "response size limit reached"],
  lineage_nodes: ["limite de cellules atteinte", "cell limit reached"],
  lineage_depth: ["profondeur maximale atteinte", "maximum depth reached"],
  depth: ["limite de profondeur atteinte", "depth limit reached"],
  node_budget: ["limite de nœuds de lignage atteinte", "lineage node limit reached"],
  edge_budget: ["limite de relations de lignage atteinte", "lineage relationship limit reached"],
  workbook_budget: ["limite de classeurs de lignage atteinte", "lineage workbook limit reached"],
  time_budget: ["durée maximale du lignage atteinte", "lineage time limit reached"],
  link_budget: ["limite de liens atteinte", "link limit reached"],
  link_build_time: ["durée de construction des liens atteinte", "link build time limit reached"],
  source_index_incomplete: ["index source incomplet", "source index incomplete"],
  external_facts_rebuild_required: ["reconstruction des liens externes requise", "external-link rebuild required"],
};
const SEARCH_BUDGET_LABELS = {
  limit: ["cellules", "cells"],
  nodes: ["cellules", "cells"],
  edges: ["relations", "relationships"],
  workbooks: ["classeurs", "workbooks"],
  depth: ["profondeur", "depth"],
  milliseconds: ["millisecondes", "milliseconds"],
  responseBytes: ["octets de réponse", "response bytes"],
  readerSamples: ["liens lecteurs détaillés", "reader links detailed"],
  unresolvedReaderSamples: ["liens lecteurs non résolus détaillés", "unresolved reader links detailed"],
  connections: ["classeurs ouverts simultanément", "workbooks open at once"],
};

function searchText(value, english) {
  if (Array.isArray(value)) return value[english ? 1 : 0];
  if (value && typeof value === "object") return JSON.stringify(value);
  return value == null ? "" : String(value);
}

function provenanceLabel(value, english) {
  const labels = {
    saved_input: ["valeur saisie", "saved input"],
    saved_cache: ["cache enregistré · non recalculé", "saved cache · not recalculated"],
    calculated: ["valeur calculée", "calculated value"],
    formula: ["résultat de formule enregistré", "saved formula result"],
  };
  return searchText(labels[value] || ["origine enregistrée", "saved provenance"], english);
}

function LibrarySearch({ apiBase = "api/", language = "fr", onClose }) {
  const english = language === "en";
  const text = (fr, en) => (english ? en : fr);
  const [query, setQuery] = useState("");
  const [intent, setIntent] = useState("search");
  const [askRun, setAskRun] = useState(null);
  const [askEpoch, setAskEpoch] = useState(0);
  const [askQuestion, setAskQuestion] = useState("");
  const [askConfig, setAskConfig] = useState(null);
  const [aiConfigFailed, setAiConfigFailed] = useState(false);
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
  const [statusScope, setStatusScope] = useState(null);
  const [catalogPollTimedOut, setCatalogPollTimedOut] = useState(false);
  const [links, setLinks] = useState([]);
  const [linksLoading, setLinksLoading] = useState(false);
  const [linksPagination, setLinksPagination] = useState(null);
  const [linksCoverage, setLinksCoverage] = useState(null);
  const [linkStatus, setLinkStatus] = useState("");
  const [pinTargets, setPinTargets] = useState({});
  const [traceEvents, setTraceEvents] = useState([]);
  const [pending, setPending] = useState({ search: null, lineage: null, upload: false });
  const busy = Object.values(pending).some(Boolean);
  const [error, setError] = useState("");
  const [attachment, setAttachment] = useState(null);
  const [recoverableAttachments, setRecoverableAttachments] = useState([]);
  const [uploadProgress, setUploadProgress] = useState("");
  const [occurrences, setOccurrences] = useState({});
  const [visibleResultCount, setVisibleResultCount] = useState(10);
  const [expandedRevision, setExpandedRevision] = useState(null);
  const traceSummary = useRef(null);
  const controllers = useRef({ search: null, lineage: null });
  const canvas = useRef(null);
  const flow = useRef(null);
  const searchRequestId = useRef(0);
  const lineageRequestId = useRef(0);
  const statusRequestId = useRef(0);
  const linksRequestId = useRef(0);
  const traceId = useRef(0);
  const restoredAttachments = useRef(false);
  const currentStatusScope = useRef(attachment?.scopeId || null);
  currentStatusScope.current = attachment?.scopeId || null;
  const currentLinkScope = useRef(attachment?.scopeId || null);
  currentLinkScope.current = attachment?.scopeId || null;

  useEffect(() => {
    const summary = traceSummary.current, main = summary.closest("main");
    const observer = new ResizeObserver(() => {
      const height = summary.getBoundingClientRect().height;
      if (height) main.style.setProperty("--library-trace-height", `${height}px`);
    });
    observer.observe(summary);
    return () => {
      observer.disconnect();
      main.style.removeProperty("--library-trace-height");
      Object.values(controllers.current).forEach(controller => controller?.abort());
    };
  }, []);

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
  useEffect(() => {
    let active = true;
    request("library/ask/config")
      .then(value => { if (active) setAskConfig(value); })
      .catch(() => { if (active) setAiConfigFailed(true); });
    return () => { active = false; };
  }, []);
  const showIntent = (value) => {
    invalidateAnswer();
    if (value === "ask") setAskQuestion(query);
    setIntent(value);
    if (value === "search") requestAnimationFrame(() => document.getElementById("library-search-query")?.focus());
  };
  const scopedParams = (scope = attachment?.scopeId) => {
    const params = new URLSearchParams();
    if (scope) params.set("scope", scope);
    return params;
  };
  const refreshStatus = async (scope = attachment?.scopeId) => {
    const requestId = ++statusRequestId.current;
    const params = scopedParams(scope);
    const value = await request(`library/status${params.size ? `?${params}` : ""}`);
    if (
      requestId !== statusRequestId.current
      || (scope || null) !== currentStatusScope.current
    ) return value;
    setStatus(value);
    setStatusScope(scope || null);
    return value;
  };
  const refreshLinks = async (scope = attachment?.scopeId, statusFilter = linkStatus, offset = 0) => {
    const requestId = ++linksRequestId.current;
    const scopeId = scope || null;
    const params = scopedParams(scopeId);
    params.set("limit", "50");
    params.set("offset", String(offset));
    if (statusFilter) params.set("status", statusFilter);
    setLinksLoading(true);
    try {
      const value = await request(`library/links?${params}`);
      if (requestId !== linksRequestId.current || scopeId !== currentLinkScope.current) return;
      setLinks((current) => offset ? [...current, ...(value.links || [])] : value.links || []);
      setLinksPagination(value.pagination || null);
      setLinksCoverage(value.coverage || null);
    } finally {
      if (requestId === linksRequestId.current) setLinksLoading(false);
    }
  };
  useEffect(() => {
    let active = true;
    Promise.all([request("projects"), request("library/attachments")])
      .then(([projectData, attachmentData]) => {
        if (!active) return;
        setProjects(projectData.projects || []);
        const available = attachmentData.attachments || [];
        setRecoverableAttachments(available);
        if (!restoredAttachments.current && available.length) {
          restoredAttachments.current = true;
          const recover = available[0];
          setAttachment({ project: recover.project, scopeId: recover.scopeId });
          setProjectId(recover.project.id);
          setMode("attachment");
        }
      })
      .catch((cause) => active && setError(cause.message));
    return () => { active = false; };
  }, []);

  useEffect(() => {
    let active = true;
    refreshStatus().catch((cause) => {
      if (active) setError(cause.message);
    });
    return () => { active = false; };
  }, [attachment?.scopeId]);

  useEffect(() => {
    const scope = attachment?.scopeId || null;
    if (statusScope !== scope) {
      ++linksRequestId.current;
      setLinksLoading(false);
      setLinks([]);
      setLinksPagination(null);
      setLinksCoverage(null);
      return undefined;
    }
    const state = status?.buildState?.status;
    if (state === "queued" || state === "building") {
      ++linksRequestId.current;
      setLinksLoading(false);
      setLinks([]);
      setLinksPagination(null);
      setLinksCoverage(null);
      return undefined;
    }
    let active = true;
    refreshLinks(scope).catch((cause) => {
      if (active) setError(cause.message);
    });
    return () => { active = false; };
  }, [attachment?.scopeId, linkStatus, statusScope, status?.buildState?.status]);

  useEffect(() => {
    const state = status?.buildState?.status;
    if (state !== "queued" && state !== "building" && !(state === "ready" && status?.stale)) {
      setCatalogPollTimedOut(false);
      return undefined;
    }
    let active = true;
    let attempts = 0;
    let timer;
    const poll = async () => {
      if (!active) return;
      attempts += 1;
      try {
        await refreshStatus();
      } catch (cause) {
        if (active) setError(cause.message);
      }
      if (!active) return;
      if (attempts >= 130) {
        setCatalogPollTimedOut(true);
        return;
      }
      timer = setTimeout(poll, 1000);
    };
    timer = setTimeout(poll, 1000);
    return () => {
      active = false;
      clearTimeout(timer);
    };
  }, [status?.buildState?.status, status?.stale, attachment?.scopeId]);

  const addTrace = (type, trace, detail = "") => {
    setTraceEvents((events) => [...events, { type, trace, detail, status: "completed", at: new Date().toISOString() }]);
  };
  const supersedeRequests = (types) => {
    for (const type of types) {
      if (type === "search") ++searchRequestId.current;
      if (type === "lineage") ++lineageRequestId.current;
      controllers.current[type]?.abort();
      controllers.current[type] = null;
    }
    setPending((current) => ({ ...current, ...Object.fromEntries(types.map(type => [type, null])) }));
    setTraceEvents(events => events.map(event => types.includes(event.type) && event.status === "pending"
      ? { ...event, status: "superseded" } : event));
  };
  const invalidateAnswer = (clearSearch = true) => {
    setAskRun(null);
    setAskEpoch(value => value + 1);
    supersedeRequests(["search", "lineage"]);
    setLineage(null);
    setExpandedRevision(null);
    if (clearSearch) {
      setSearch(null);
      setSelected(null);
      setOccurrences({});
    }
  };
  const apiScope = (selectedMode = mode, selectedProject = projectId, selectedRevision = revisionId) => {
    const params = scopedParams();
    if (selectedMode === "project" && selectedProject) params.set("project", selectedProject);
    if (selectedMode === "attachment" && attachment?.project?.id) params.set("project", attachment.project.id);
    if (selectedMode === "revision" && selectedRevision) {
      const separator = selectedRevision.indexOf(":");
      params.set("revision", selectedRevision);
      if (separator > 0) params.set("project", selectedRevision.slice(0, separator));
    }
    return params;
  };
  const performSearch = async (offset = 0, generation = "") => {
    const appendResults = offset > 0 && search?.generation === generation;
    supersedeRequests(appendResults ? ["search"] : ["search", "lineage"]);
    const requestId = ++searchRequestId.current;
    const controller = new AbortController();
    controllers.current.search = controller;
    setPending(current => ({ ...current, search: requestId }));
    setError("");
    if (!appendResults) {
      setSearch(null);
      setLineage(null);
      setSelected(null);
      setVisibleResultCount(10);
      setExpandedRevision(null);
    }
    const traceRequestId = ++traceId.current;
    setTraceEvents((events) => [...events, {
      type: "search",
      status: "pending",
      trace: null,
      detail: query,
      queryText: query,
      scopeLabel: mode === "project"
        ? projects.find((project) => project.id === projectId)?.filename || text("classeur choisi", "selected workbook")
        : mode === "revision"
          ? revisionLabel(revisions.find((item) => item.revision?.id === revisionId) || {})
          : mode === "attachment"
            ? attachment?.project?.filename || text("fichier temporaire", "temporary file")
            : text("toute la bibliothèque", "whole library"),
      requestId: traceRequestId,
      at: new Date().toISOString(),
    }]);
    try {
      const params = apiScope();
      params.set("q", query);
      params.set("kinds", kinds.join(","));
      params.set("limit", "50");
      params.set("offset", String(offset));
      if (generation) params.set("generation", generation);
      const result = await request(`library/search?${params}`, { signal: controller.signal });
      if (requestId !== searchRequestId.current) return;
      setSearch((previous) => appendResults && previous?.generation === generation
        ? { ...result, results: [...previous.results, ...(result.results || [])] }
        : result);
      if (!appendResults) setOccurrences({});
      setTraceEvents((events) => events.map((event) =>
        event.requestId === traceRequestId
          ? { ...event, status: "completed", trace: result.trace, detail: `${result.results?.length || 0}`, complete: result.coverage?.complete }
          : event,
      ));
    } catch (cause) {
      if (requestId === searchRequestId.current) {
        const message = friendlyError(cause);
        setSearch(null);
        setError(message);
        setTraceEvents((events) => events.map((event) =>
          event.requestId === traceRequestId ? { ...event, status: "failed", error: message, technicalError: cause.message } : event,
        ));
      }
    } finally {
      setPending(current => current.search === requestId ? { ...current, search: null } : current);
    }
  };
  const selectResult = async (hit, requestedDirection = direction, scope = attachment?.scopeId) => {
    supersedeRequests(["lineage"]);
    const requestId = ++lineageRequestId.current;
    setSelected(hit);
    setLineage(null);
    if (!hit.handle) return;
    const controller = new AbortController();
    controllers.current.lineage = controller;
    setPending(current => ({ ...current, lineage: requestId }));
    setError("");
    const traceRequestId = ++traceId.current;
    setTraceEvents((events) => [...events, {
      type: "lineage", status: "pending", trace: null, detail: requestedDirection,
      requestId: traceRequestId, at: new Date().toISOString(),
    }]);
    try {
      const params = scopedParams(scope);
      params.set("handle", hit.handle);
      params.set("direction", requestedDirection);
      params.set("depth", "20");
      params.set("limit", "200");
      const result = await request(`library/lineage?${params}`, { signal: controller.signal });
      if (requestId !== lineageRequestId.current) return;
      setLineage(result);
      setTraceEvents((events) => events.map((event) =>
        event.requestId === traceRequestId
          ? { ...event, status: "completed", trace: result.trace, complete: result.coverage?.complete }
          : event,
      ));
    } catch (cause) {
      if (requestId === lineageRequestId.current) {
        const message = friendlyError(cause);
        setLineage(null);
        setError(message);
        setTraceEvents((events) => events.map((event) =>
          event.requestId === traceRequestId ? { ...event, status: "failed", error: message, technicalError: cause.message } : event,
        ));
      }
    } finally {
      setPending(current => current.lineage === requestId ? { ...current, lineage: null } : current);
    }
  };
  const graphView = lineage ? revisionLineageView(lineage, expandedRevision) : null;
  const boundaryText = (boundaries) => [...new Set(boundaries.map(boundary => boundary.stalePin
    ? text("Épingle périmée : choisissez à nouveau la cible", "Stale pin: choose the target again")
    : searchText(SEARCH_BOUNDARY_LABELS[boundary.reason] || ["Source à vérifier", "Source to verify"], english)))].join(" · ");
  useEffect(() => {
    if (!canvas.current || !lineage) return undefined;
    const view = revisionLineageView(lineage, expandedRevision);
    const instance = mount(canvas.current, {
      compact: true,
      stack: !view.expandedRevision,
      onSelect: (node) => {
        if (node.collapsed) setExpandedRevision(node.id);
        else {
          const found = (lineage.nodes || []).find(item => item.id === node.id);
            if (found && intent === "search") selectResult(found, direction);
        }
      },
      relationshipStatus: (model) => {
        const drawn = model.edges.filter(edge => !edge.hidden);
        const represented = drawn.reduce((count, edge) => count + (edge.weight || 1), 0);
        return `${countLabel(drawn.length, english, ["liaison affichée", "liaisons affichées"], ["link drawn", "links drawn"])} · ${countLabel(represented, english, ["dépendance", "dépendances"], ["dependency", "dependencies"])} · ${countLabel(0, english, ["lien synthétique", "liens synthétiques"], ["synthetic link", "synthetic links"])}`;
      },
    });
    flow.current = instance;
    const nodes = view.nodes.map(node => {
      const revision = node.revision || node.workbookRevision || {};
      const stops = boundaryText(node.boundaries || []);
      const stopped = (node.boundaries || []).some(boundary => boundary.reason !== "input_constant");
      if (node.collapsed) return {
        ...node, type: "sheet", kindLabel: `${text("Révision", "Revision")} ${revision.projectId?.slice(0, 8) || ""} · ${revision.sha256?.slice(0, 8) || ""}`,
        title: node.filename || revision.projectId,
        subtitle: `${countLabel(node.members.length, english, ["cellule", "cellules"], ["cell", "cells"])} · ${countLabel(node.internalEdges, english, ["relation interne", "relations internes"], ["internal relationship", "internal relationships"])}`,
        valueLabel: stops || text("Développer pour suivre les cellules", "Expand to follow cells"),
        count: node.members.length, tone: stopped ? "boundary" : "", original: node,
      };
      return {
        ...node, nodeId: node.id,
        kindLabel: kindLabel(node.kind), title: `${node.sheet}!${node.address || node.cell || ""}`,
        subtitle: stops || provenanceLabel(node.valueProvenance, english),
        valueLabel: node.value == null ? text("Aucune valeur enregistrée", "No saved value") : searchText(node.value, english),
        tone: stopped ? "boundary" : node.boundaries.length ? "origin" : "",
        original: node,
      };
    });
    const edges = view.edges.map(edge => ({ ...edge, label: edge.kind === "external_link"
      ? text("lien externe", "external link") : edge.kind === "range"
        ? text("plage", "range") : text("formule", "formula") }));
    const frame = requestAnimationFrame(() => {
      if (flow.current === instance) instance.render(nodes, edges, { language: english ? "en" : "fr" });
    });
    return () => { cancelAnimationFrame(frame); instance.destroy(); flow.current = null; };
  }, [lineage, expandedRevision, english, intent]);

  const toggleKind = (kind) => {
    invalidateAnswer();
    setKinds((selectedKinds) => selectedKinds.includes(kind)
      ? selectedKinds.filter((item) => item !== kind)
      : [...selectedKinds, kind]);
  };
  const changeScope = (update, value) => {
    invalidateAnswer();
    setError("");
    update(value);
  };
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
          rows: [...currentPage.rows, ...(page.results || []).filter((row) =>
            !currentPage.rows.some((existing) => existing.address === row.address && (existing.sheet || hit.sheet) === (row.sheet || hit.sheet)),
          )],
          after: page.pagination?.nextAfter ?? currentPage.after,
          done: page.pagination?.nextAfter == null,
        },
      }));
      addTrace("occurrences", { generation: group.generation, capsHit: page.coverage?.capsHit || [], complete: page.coverage?.complete }, `${page.results?.length || 0}`);
    } catch (cause) {
      setError(friendlyError(cause));
      addTrace("occurrences", null, friendlyError(cause));
    }
  };
  const attachFile = async (file) => {
    const previous = attachment;
    invalidateAnswer();
    setPending(current => ({ ...current, upload: true }));
    setError("");
    let uploadId = "";
    let scopeId = "";
    const traceRequestId = ++traceId.current;
    setTraceEvents((events) => [...events, {
      type: "upload", status: "pending", trace: null, detail: text("Import en préparation", "Preparing import"),
      requestId: traceRequestId, at: new Date().toISOString(),
    }]);
    try {
      if (previous) {
        await request(`library/attachments/${previous.scopeId}`, { method: "DELETE", body: "{}" });
        setAttachment(null);
        setRecoverableAttachments((items) => items.filter((item) => item.scopeId !== previous.scopeId));
      }
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
      scopeId = imported.project.searchScopeId || uploadId;
      let task = imported.task;
      let polls = 0;
      while (task.status !== "succeeded" && task.status !== "failed" && task.status !== "cancelled" && polls < 1200) {
        setUploadProgress(`${text("Import", "Import")} · ${task.progress?.phase || task.status}`);
        setTraceEvents((events) => events.map((event) =>
          event.requestId === traceRequestId
            ? { ...event, detail: `${text("Import", "Import")} · ${task.progress?.phase || task.status}` }
            : event,
        ));
        await new Promise((resolve) => setTimeout(resolve, 350));
        task = (await request(`tasks/${task.id}`)).task;
        polls += 1;
      }
      if (polls >= 1200 && task.status !== "succeeded") {
        throw new Error(text("Délai d’import dépassé. Le fichier temporaire peut être retiré ci-dessous.", "Import timed out. The temporary file can be removed below."));
      }
      if (task.status !== "succeeded") throw new Error(task.error?.message || text("Import échoué.", "Import failed."));
      setAttachment({ project: imported.project, scopeId });
      setProjectId(imported.project.id);
      setMode("attachment");
      setUploadProgress(text("Fichier joint à cette recherche uniquement.", "File attached to this search only."));
      setRecoverableAttachments((items) => [{ project: imported.project, scopeId, expiresAt: imported.project.searchScopeExpiresAt }, ...items.filter((item) => item.scopeId !== scopeId)]);
      setTraceEvents((events) => events.map((event) =>
        event.requestId === traceRequestId
          ? { ...event, status: "completed", detail: text("Import terminé · fichier temporaire récupérable pendant 24 h", "Import complete · temporary file recoverable for 24 hours") }
          : event,
      ));
      await refreshStatus(scopeId);
    } catch (cause) {
      let cleanupError = "";
      if (scopeId) {
        try {
          await request(`library/attachments/${scopeId}`, { method: "DELETE", body: "{}" });
        } catch {
          cleanupError = text("Le nettoyage a échoué ; récupérez ou retirez le fichier dans la liste des fichiers temporaires.", "Cleanup failed; recover or remove the file in the temporary files list.");
          try {
            const data = await request("library/attachments");
            setRecoverableAttachments(data.attachments || []);
          } catch (listError) {
            cleanupError = `${cleanupError} ${listError.message}`;
          }
        }
      } else if (uploadId) {
        try { await request(`uploads/${uploadId}`, { method: "DELETE", body: "{}" }); } catch { /* Preserve the original import failure. */ }
      }
      setTraceEvents((events) => events.map((event) =>
        event.requestId === traceRequestId ? { ...event, status: "failed", error: friendlyError(cause), technicalError: cause.message } : event,
      ));
      setError([friendlyError(cause), cleanupError].filter(Boolean).join(" · "));
    } finally {
      setPending(current => ({ ...current, upload: false }));
    }
  };
  const keepAttachment = async () => {
    if (!attachment) return;
    const current = attachment;
    const root = selected;
    invalidateAnswer(false);
    const mutationGeneration = searchRequestId.current;
    setSearch(null);
    setError("");
    try {
      await request(`library/attachments/${current.scopeId}/keep`, { method: "POST", body: "{}" });
      setAttachment(null);
      setRecoverableAttachments((items) => items.filter((item) => item.scopeId !== current.scopeId));
      setMode("project");
      setUploadProgress(text("Classeur conservé dans la bibliothèque.", "Workbook kept in the library."));
      const data = await request("projects");
      setProjects(data.projects || []);
      await Promise.all([refreshStatus(null), refreshLinks(null)]);
      if (root && mutationGeneration === searchRequestId.current) await selectResult(root, direction, null);
    } catch (cause) {
      setError(cause.message);
    }
  };
  const removeAttachment = async () => {
    if (!attachment) return;
    const current = attachment;
    invalidateAnswer();
    setError("");
    try {
      await request(`library/attachments/${current.scopeId}`, { method: "DELETE", body: "{}" });
      setAttachment(null);
      setRecoverableAttachments((items) => items.filter((item) => item.scopeId !== current.scopeId));
      setMode("library");
      setUploadProgress("");
      await Promise.all([refreshStatus(null), refreshLinks(null)]);
    } catch (cause) {
      setError(cause.message);
    }
  };
  const pin = async (link, remove = false) => {
    const root = selected;
    invalidateAnswer(false);
    const mutationGeneration = searchRequestId.current;
    setError("");
    try {
      const params = scopedParams();
      if (remove) {
        await request(`library/pins/${encodeURIComponent(link.id)}${params.size ? `?${params}` : ""}`, { method: "DELETE", body: "{}" });
      } else {
        const currentPin = link.pin && `${link.pin.projectId}:${link.pin.sha256}`;
        const [targetProject, targetSha256] = (pinTargets[link.id] || currentPin || "").split(":");
        if (!targetProject || !targetSha256) throw new Error(text("Choisissez une révision cible.", "Choose a target revision."));
        await request("library/pins", {
          method: "POST",
          body: JSON.stringify({ linkId: link.id, targetProject, targetSha256, ...(attachment ? { scope: attachment.scopeId } : {}) }),
        });
      }
      await refreshLinks();
      await refreshStatus();
      if (root && mutationGeneration === searchRequestId.current) await selectResult(root, direction);
    } catch (cause) {
      setError(cause.message);
    }
  };
  const results = search?.results || [];
  const revisions = status?.revisions || [];
  const complete = search?.coverage?.complete;
  const caps = search?.coverage?.capsHit || search?.trace?.capsHit || [];
  const linksPending = linksLoading
    || statusScope !== (attachment?.scopeId || null)
    || status?.buildState?.status === "queued"
    || status?.buildState?.status === "building";
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
      searchText(SEARCH_KIND_LABELS[node.kind] || node.kind, english),
    ].filter(Boolean).join(" · "),
  ]));
  const traceTime = (value) => value == null ? "" : `${value} ms`;
  const revisionLabel = (revision) => {
    const sha = revision?.sha256 || "";
    return `${revision?.filename || ""} · ${revision?.projectId?.slice(0, 8) || ""} · ${sha.slice(0, 8)}`;
  };
  const kindLabel = (kind) => searchText(SEARCH_KIND_LABELS[kind] || kind, english);
  const roleLabel = (role) => searchText(SEARCH_ROLE_LABELS[role] || role || text("étiquette", "label"), english);
  const statusLabel = (value) => searchText(SEARCH_STATUS_LABELS[value] || value, english);
  const capLabel = (value) => searchText(SEARCH_CAP_LABELS[value] || ["limite de traitement atteinte", "processing limit reached"], english);
  const friendlyError = (cause) => {
    const message = String(cause?.message || cause || "");
    if (message.includes("query_chars")) return text("La requête est trop longue. Raccourcissez-la puis réessayez.", "The query is too long. Shorten it and try again.");
    if (message.includes("Catalog building")) return text("L’indexation est en cours. Réessayez dans quelques instants.", "The index is building. Try again shortly.");
    if (message.includes("Library revision or cell not found")) return text("Cette révision ou cette cellule n’est plus disponible.", "This revision or cell is no longer available.");
    if (message.includes("Library build failed")) return text("L’indexation a échoué. Vérifiez le stockage et réimportez un classeur.", "Indexing failed. Check storage and reimport a workbook.");
    if (message.toLowerCase().includes("not a zip file")) return text("Le fichier Excel est endommagé ou incomplet.", "The Excel file is damaged or incomplete.");
    return message;
  };
  const latestTraceEvent = traceEvents[traceEvents.length - 1];
  const latestTraceSummary = !latestTraceEvent
    ? text("Aucune requête effectuée. La trace apparaîtra ici, même sans résultat.", "No request yet. The trace will appear here, including when no results are found.")
    : latestTraceEvent.status === "pending"
      ? (latestTraceEvent.type === "search" ? text("Recherche en cours…", "Searching…") : latestTraceEvent.type === "lineage" ? text("Lignage en cours…", "Tracing lineage…") : latestTraceEvent.detail)
    : latestTraceEvent.status === "superseded"
      ? text("Requête remplacée · lancez la recherche dans le périmètre choisi.", "Request superseded · search the selected scope.")
    : latestTraceEvent.error
      ? `${latestTraceEvent.type === "search" ? text("Échec de la recherche", "Search failed") : latestTraceEvent.type === "lineage" ? text("Échec du lignage", "Lineage failed") : text("Échec de l’import", "Import failed")} · ${friendlyError(latestTraceEvent.error)}`
        : latestTraceEvent.type === "lineage"
          ? `${latestTraceEvent.detail === "downstream" ? text("En aval", "Downstream") : text("En amont", "Upstream")} · ${latestTraceEvent.complete === false ? text("parcours incomplet ou limité", "incomplete or capped traversal") : text("couverture indiquée près du graphe", "coverage shown beside the graph")}`
      : (latestTraceEvent.complete ?? latestTraceEvent.trace?.complete) === false
        ? text("Résultat incomplet ou limité · détails de la trace ci-dessous", "Incomplete or capped result · trace details below")
        : latestTraceEvent.type === "search"
          ? text("Recherche terminée · consultez la couverture des résultats", "Search complete · see results coverage")
          : latestTraceEvent.detail;
  const catalogFailureReason = status?.buildState?.errorType === "OSError"
    ? text("Le stockage disponible est insuffisant pour créer l’index.", "There is not enough storage to build the index.")
    : status?.buildState?.errorType
      ? text("Le classeur n’a pas pu être indexé.", "The workbook could not be indexed.")
      : "";
  const lineageGroups = [...lineageNodes.reduce((groups, node) => {
    const revision = node.revision || node.workbookRevision || {};
    const key = revision.id || `${node.projectId || ""}:${revision.sha256 || ""}`;
    if (!groups.has(key)) groups.set(key, { key, title: revisionLabel({ filename: node.filename || node.projectName, projectId: revision.projectId || node.projectId, sha256: revision.sha256 || node.sha256 }), nodes: [] });
    groups.get(key).nodes.push(node);
    return groups;
  }, new Map()).values()];
  return (
    <div className="library-search-view" onFocusCapture={(event) => {
      const bar = traceSummary.current, target = event.target;
      if (!bar || bar.contains(target)) return;
      const bottom = bar.getBoundingClientRect().bottom + 8;
      const top = target.getBoundingClientRect().top;
      if (top < bottom) bar.closest("main").scrollBy({ top: top - bottom });
    }}>
      <header className="library-search-heading">
        <div className="library-search-title">
          <button type="button" className="library-back" onClick={onClose}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" aria-hidden="true"><path d="m12 5-7 7 7 7M5 12h14" /></svg>
            {text("Retour au classeur", "Back to workbook")}
          </button>
          <h2>{text("Rechercher", "Search")}</h2>
        </div>
        <p>{text("Retrouvez une valeur, une cellule ou un libellé dans les classeurs importés.", "Find a value, cell or label in imported workbooks.")}</p>
      </header>
      <section ref={traceSummary} className="library-trace-summary" aria-live="polite">
        <strong>{text("Dernière opération", "Latest operation")}</strong>
        <span>{intent === "ask" ? askRun?.status === "running" ? text("Question en cours · trace disponible", "Question running · trace available") : askRun ? text("Question terminée · consultez la trace", "Question finished · see trace") : text("Aucune question exécutée.", "No question run yet.") : latestTraceSummary}</span>
        <a href="#library-trace-heading" onClick={(event) => {
          event.preventDefault();
          document.getElementById("library-trace")?.setAttribute("open", "");
          requestAnimationFrame(() => {
            const heading = document.getElementById("library-trace-heading");
            heading?.focus({ preventScroll: true });
            heading?.scrollIntoView({ block: "start" });
          });
        }}>{text("Trace détaillée", "Detailed trace")}</a>
      </section>
      <section className="library-search-controls" aria-label={text("Recherche dans la bibliothèque", "Search the library")}>
        {intent === "ask" && <>
          <button type="button" className="library-search-return" onClick={() => showIntent("search")}>{text("Retour à la recherche", "Back to search")}</button>
          <AskQuestion key={`${apiScope().toString()}:${askEpoch}`} config={askConfig} initialQuestion={query} draftQuestion={askQuestion} onQuestionChange={setAskQuestion} sourceOptions={revisions} sourceProjectId={apiScope().get("project") || ""} sourceLocked={mode === "revision" || mode === "attachment"} onSourceChange={(value) => changeScope(() => { setProjectId(value); setMode(value ? "project" : "library"); })} request={request} scope={Object.fromEntries(apiScope())} english={english} onRun={setAskRun} onClear={() => { setSearch(null); setSelected(null); setLineage(null); }} onEvidence={(answer, index = 0) => { setSelected(answer.cells[index]); setLineage(answer.lineages[index]); setDirection("upstream"); setExpandedRevision(null); }} />
        </>}
        {intent === "search" && <form onSubmit={(event) => { event.preventDefault(); performSearch(); }}>
          <label htmlFor="library-search-query">{text("Valeur, cellule ou libellé", "Value, cell or label")}</label>
          <div className="library-search-form-row">
            <input id="library-search-query" type="search" placeholder={text("Ex. : chiffre d’affaires, A12, 42", "E.g. revenue, A12, 42")} aria-describedby="library-search-help" value={query} onChange={(event) => { setQuery(event.target.value); invalidateAnswer(); setError(""); }} />
            <button type="submit" className="primary" disabled={busy}>{text("Rechercher", "Search")}</button>
            {askConfig?.available && <button type="button" className="library-ask-entry" onClick={() => showIntent("ask")}>{text("Poser une question à l’IA", "Ask AI a question")}</button>}
          </div>
          <p id="library-search-help" className="help">{text("Recherche sans IA · tous les termes doivent correspondre.", "Search without AI · all terms must match.")}</p>
          {aiConfigFailed && <p className="help" role="status">{text("La configuration IA est indisponible. Vous pouvez continuer la recherche sans IA.", "AI configuration is unavailable. You can continue searching without AI.")}</p>}
        </form>}
        <details className="library-advanced">
          <summary>{text("Recherche avancée", "Advanced search")}<span>{mode === "attachment" ? attachment?.project.filename : mode === "project" ? projects.find(project => project.id === projectId)?.filename || text("Choisir un classeur", "Choose a workbook") : mode === "revision" ? text("Une révision", "One revision") : text("Toute la bibliothèque", "Whole library")}{intent === "search" && kinds.length !== 5 ? ` · ${kinds.length} ${kinds.length === 1 ? text("type sélectionné", "selected kind") : text("types sélectionnés", "selected kinds")}` : ""}</span></summary>
          <div className="library-advanced-content">
        <fieldset className="library-scope">
          <legend>{text("Portée", "Scope")}</legend>
          {[
            ["library", text("Toute la bibliothèque", "Whole library")],
            ["project", text("Classeurs", "Workbooks")],
            ["revision", text("Révision", "Revision")],
            ["attachment", text("Fichier joint", "Attached file")],
          ].map(([value, label]) => (
            <button key={value} type="button" className={mode === value ? "is-active" : ""} aria-pressed={mode === value} disabled={value === "attachment" && !attachment} onClick={() => changeScope(setMode, value)}>{label}</button>
          ))}
          {mode === "project" && <label>{text("Classeur", "Workbook")}<select value={projectId} onChange={(event) => changeScope(setProjectId, event.target.value)}><option value="">{text("Choisir…", "Choose…")}</option>{projects.map((project) => <option key={project.id} value={project.id}>{project.filename} · {project.id.slice(0, 8)}</option>)}</select></label>}
          {mode === "revision" && <label>{text("Révision", "Revision")}<select value={revisionId} onChange={(event) => changeScope(setRevisionId, event.target.value)}><option value="">{text("Choisir…", "Choose…")}</option>{revisions.map((revision) => <option key={revision.revision?.id || revision.projectId} value={revision.revision?.id}>{revisionLabel({ filename: revision.filename, projectId: revision.projectId, sha256: revision.revision?.sha256 })}</option>)}</select></label>}
          {mode === "attachment" && attachment && <span>{attachment.project.filename} · {attachment.project.id.slice(0, 8)}</span>}
        </fieldset>
        {intent === "search" && <fieldset className="library-kind-filter">
          <legend>{text("Types", "Kinds")}</legend>
          {["cell", "sheet", "name", "table", "column"].map((kind) => <label key={kind}><input type="checkbox" checked={kinds.includes(kind)} onChange={() => toggleKind(kind)} />{kindLabel(kind)}</label>)}
        </fieldset>}
        <div className="library-attach-control">
          <label htmlFor="library-attach-file">{text("Joindre un classeur à cette recherche", "Attach a workbook to this search")}</label>
          <input id="library-attach-file" className="library-file-input" type="file" accept=".xlsx,.xlsm,.xlam,.xla" disabled={busy} onChange={(event) => { const file = event.target.files?.[0]; if (file) attachFile(file); event.target.value = ""; }} />
          <label className="library-file-picker" htmlFor="library-attach-file"><span>{text("Choisir un fichier Excel", "Choose an Excel file")}</span><span>{attachment?.project?.filename || text("Aucun fichier sélectionné", "No file selected")}</span></label>
          {uploadProgress && <p role="status">{uploadProgress}</p>}
          {attachment && <div className="library-attachment-actions"><button type="button" onClick={keepAttachment}>{text("Conserver dans la bibliothèque", "Keep in library")}</button><button type="button" onClick={removeAttachment}>{text("Retirer le fichier joint", "Remove attached file")}</button></div>}
          {recoverableAttachments.length > 0 && <div className="library-recoverable-attachments"><strong>{text("Fichiers temporaires récupérables (expiration après 24 h)", "Recoverable temporary files (expire after 24 h)")}</strong>{recoverableAttachments.map((item) => <div key={item.scopeId}><button type="button" onClick={() => changeScope(() => { setAttachment({ project: item.project, scopeId: item.scopeId }); setProjectId(item.project.id); setMode("attachment"); })}>{text("Récupérer", "Recover")} · {item.project.filename} · {item.project.id.slice(0, 8)}</button>{attachment?.scopeId !== item.scopeId && <button type="button" onClick={async () => { try { await request(`library/attachments/${item.scopeId}`, { method: "DELETE", body: "{}" }); setRecoverableAttachments((items) => items.filter((entry) => entry.scopeId !== item.scopeId)); } catch (cause) { setError(cause.message); } }}>{text("Supprimer", "Remove")}</button>}</div>)}</div>}
        </div>
          </div>
        </details>
      </section>
      <section className="library-catalog-state" aria-live="polite">
        <strong>{text("Catalogue", "Catalog")}</strong>
        <span>{({
          queued: text("Mise à jour en attente", "Update queued"),
          building: text("Mise à jour en cours", "Update in progress"),
          failed: text("Mise à jour échouée", "Update failed"),
          ready: status?.stale || status?.pending ? text("Mise à jour en attente", "Update pending") : text("À jour", "Up to date"),
        })[status?.buildState?.status] || text("État inconnu", "Unknown state")}</span>
        {status?.stale && <span>{text("Les résultats peuvent être incomplets pendant la mise à jour.", "Results may be incomplete while the index updates.")}</span>}
        {status?.buildState?.status === "failed" && <span>{catalogFailureReason} {text("Vérifiez l’espace disponible.", "Check available storage.")}</span>}
        {catalogPollTimedOut && <span>{text("La mise à jour prend trop de temps.", "The update is taking too long.")}</span>}
        {(status?.buildState?.status === "failed" || catalogPollTimedOut) && <button type="button" onClick={() => document.getElementById("library-attach-file")?.click()}>{text("Choisir un classeur pour reconstruire l’index", "Choose a workbook to rebuild the index")}</button>}
      </section>
      {error && <p className="notice error" role="alert">{error}</p>}

      <div className="library-search-columns">
        <section className="library-results" hidden={intent === "ask"} aria-labelledby="library-results-heading">
          <h3 id="library-results-heading">{text("Résultats", "Results")}</h3>
          {search && <p className={`library-result-coverage ${complete ? "is-complete" : "is-incomplete"}`} role="status">
            {complete ? text("Recherche complète dans le périmètre et le contenu indexé.", "Search complete for the indexed content in this scope.") : text("Résultats partiels : la recherche a atteint une limite.", "Partial results: the search reached a limit.")}
            {!!caps.length && <span> {caps.map(capLabel).join(" · ")}</span>}
            {search.pagination?.hasMore && <span> {text("D’autres résultats sont disponibles.", "More results are available.")}</span>}
          </p>}
          {search && results.length === 0 && <p className="library-empty">{text("Aucun résultat affiché pour cette requête et cette portée.", "No results are displayed for this query and scope.")}</p>}
          {results.slice(0, visibleResultCount).map((hit) => {
            const labels = hit.labels || [];
            const occurrenceState = occurrences[hit.handle];
            const shownOccurrences = occurrenceState?.rows || hit.occurrences || [];
            return <article className={`library-result ${selected?.handle === hit.handle ? "is-selected" : ""}`} key={`${hit.revision?.id}:${hit.kind}:${hit.address}:${hit.sheet}:${hit.entryId || ""}`}>
              <button type="button" className="library-result-select" onClick={() => selectResult(hit)} aria-label={`${text("Afficher la piste pour", "Show lineage for")} ${hit.filename || ""} ${hit.sheet || ""}!${hit.address || ""}`}>
                <strong className="library-result-value">{searchText(hit.value, english)}</strong>
                <span>{revisionLabel({ filename: hit.filename, projectId: hit.revision?.projectId, sha256: hit.revision?.sha256 })}</span>
                <span>{hit.sheet}{hit.address ? `!${hit.address}` : ""} · {kindLabel(hit.kind)}</span>
              </button>
              {hit.formula && <p className="library-evidence">{text("Formule enregistrée", "Saved formula")}: <code>{hit.formula}</code></p>}
              {hit.valueProvenance && <p className="library-evidence">{text("Origine de la valeur", "Value source")}: {provenanceLabel(hit.valueProvenance, english)}{hit.valueProvenance === "saved_cache" && ` · ${text("aucune formule n’a été recalculée", "no formula was recalculated")}`}</p>}
              {!!labels.length && <ul className="library-labels">{labels.map((label, index) => {
                const declared = ["table_column", "defined_name"].includes(label.method) || ["table_column", "defined_name"].includes(label.role);
                const heuristic = label.certainty === "heuristic" || !declared;
                const source = [label.sourceSheet || label.sheet, label.sourceCell || label.cell].filter(Boolean).join("!");
                return <li key={`${label.text}:${label.role}:${index}`}><span>{roleLabel(label.role || label.method)}</span><strong>{label.text}</strong><small>{declared ? text("Déclaré dans le classeur", "Declared in workbook") : heuristic ? text("Inféré · heuristique", "Inferred · heuristic") : text("Source non vérifiée", "Unverified source")}{source ? ` · ${text("source", "source")} ${source}` : ""}{label.certainty ? ` · ${label.certainty === "heuristic" ? text("certitude faible", "low certainty") : text("certitude déclarée", "declared certainty")}` : ""}</small></li>;
              })}</ul>}
              {hit.occurrenceCount !== undefined && <div className="library-occurrences">
                <span>{hit.occurrenceCount} {hit.occurrenceCount === 1 ? text("occurrence", "occurrence") : text("occurrences", "occurrences")} · {hit.occurrenceCountComplete ? text("comptage complet", "complete count") : text("comptage partiel", "partial count")}</span>
                {shownOccurrences.map((occurrence) => <button key={occurrence.handle || occurrence.address} type="button" onClick={() => selectResult({ ...hit, handle: occurrence.handle, sheet: occurrence.sheet || hit.sheet, address: occurrence.address })}>{occurrence.sheet || hit.sheet}!{occurrence.address}</button>)}
                {hit.occurrenceGroup && !occurrenceState?.done && <button type="button" onClick={() => addOccurrences(hit)}>{text("Afficher la page suivante", "Show next page")}</button>}
              </div>}
            </article>;
          })}
          {results.length > visibleResultCount && <button type="button" onClick={() => setVisibleResultCount((count) => count + 10)}>{text("Afficher 10 résultats supplémentaires", "Show 10 more results")}</button>}
          {search?.pagination?.hasMore && <button type="button" onClick={() => performSearch((search.pagination.offset || 0) + search.pagination.limit, search.generation)}>{text("Page de résultats suivante", "Next results page")}</button>}
        </section>
        <section className="library-lineage" hidden={intent === "ask" && !lineage} aria-labelledby="library-lineage-heading">
          <div className="library-section-heading"><h3 id="library-lineage-heading">{text("Piste de lignage", "Lineage trail")}</h3>
            {intent === "search" && <label>{text("Sens", "Direction")}<select value={direction} onChange={(event) => { const nextDirection = event.target.value; setDirection(nextDirection); if (selected) selectResult(selected, nextDirection); }}><option value="upstream">{text("En amont", "Upstream")}</option><option value="downstream">{text("En aval", "Downstream")}</option></select></label>}
          </div>
          {selected && lineageCoverage && <p className={`library-verdict ${lineageCoverage.complete ? "is-complete" : "is-incomplete"}`} role="status">
            {lineageCoverage.exhaustive ? text("Parcours exhaustif dans le périmètre chargé.", "Traversal exhaustive within the loaded scope.") : text("Parcours non exhaustif.", "Traversal is not exhaustive.")}
            {" "}{lineageCoverage.complete ? text("Lignage complet.", "Lineage complete.") : text("Lignage incomplet.", "Lineage incomplete.")}
            {lineageCoverage.boundaries?.length > 0 && ` ${text("Arrêts", "Stops")}: ${lineageCoverage.boundaries.map((reason) => searchText(SEARCH_BOUNDARY_LABELS[reason] || reason, english)).join(", ")}.`}
            {lineageCoverage.capsHit?.length > 0 && ` ${text("Limites", "Limits")}: ${lineageCoverage.capsHit.map(capLabel).join(", ")}.`}
            {lineageBoundaries.some((boundary) => boundary.stalePin) && ` ${text("Une épingle périmée empêche de confirmer la cible.", "A stale pin prevents confirmation of the target.")}`}
          </p>}
          {lineage && <div className="library-diagnostics">
            {(lineage.edges || []).filter((edge) => edge.resolutionStatus === "resolved_by_name").length > 0 && <p>{text("Résolution par nom : ", "Matched by name: ")}{countLabel((lineage.edges || []).filter((edge) => edge.resolutionStatus === "resolved_by_name").length, english, ["relation", "relations"], ["link", "links"])}{text(" ; la révision peut différer de celle visée lors de l’enregistrement du lien.", "; the revision may differ from the one the link was saved against.")}</p>}
            {(lineage.edges || []).flatMap((edge, index) => (edge.cacheDiagnostic || []).map((diagnostic, diagnosticIndex) => <p key={`cache:${index}:${diagnosticIndex}`}>{text("Valeur mise en cache comparée à la cible actuelle", "Saved cache compared with current target")} · {diagnostic.address}: {searchText(diagnostic.cached, english)} → {searchText(diagnostic.current, english)} · {diagnostic.agreement === false ? text("différent", "different") : diagnostic.agreement === true ? text("identique", "same") : text("comparaison inconnue", "comparison unknown")}</p>))}
          </div>}
          {lineage ? <div className="library-graph-layout">
            <nav className="library-revision-navigation" aria-label={text("Classeurs du graphe et cellules regroupées", "Graph workbooks and member cells")}>
              {graphView.expandedRevision && <button type="button" onClick={() => setExpandedRevision(null)}>{text("Vue des classeurs", "Workbook overview")}</button>}
              {graphView.groups.map(group => <div key={group.id}>
                <button type="button" aria-pressed={graphView.expandedRevision === group.id} onClick={() => setExpandedRevision(group.id)}>{text("Développer", "Expand")} · {group.filename} · {group.revision.projectId?.slice(0, 8)} · {group.revision.sha256?.slice(0, 8)}</button>
                <span>{group.members.map(node => `${node.sheet}!${node.address || node.cell || ""}`).join(", ")}</span>
                {!!group.boundaries.length && <span className="library-boundary-summary">{boundaryText(group.boundaries)}</span>}
              </div>)}
            </nav>
            <p className="help library-graph-counts" role="status">{countLabel(graphView.edges.length, english, ["liaison affichée représentant", "liaisons affichées représentant"], ["link drawn representing", "links drawn representing"])} {countLabel(graphView.edges.reduce((sum, edge) => sum + (edge.weight || 1), 0), english, ["dépendance", "dépendances"], ["dependency", "dependencies"])} · {countLabel(0, english, ["lien synthétique", "liens synthétiques"], ["synthetic link", "synthetic links"])}. {countLabel(graphView.internalEdgesGrouped, english, ["relation interne regroupée", "relations internes regroupées"], ["internal relationship grouped", "internal relationships grouped"])}{text(". Développez un classeur pour suivre ses cellules. Les arrêts sont indiqués sur les cartes et dans les preuves.", ". Expand a workbook to follow its cells. Stops appear on cards and in evidence.")}</p>
            <div className="library-graph-canvas" ref={canvas} role="application" aria-label={text("Graphe de lignage. Une alternative en liste suit.", "Lineage graph. A list alternative follows.")} />
            <details className="library-accessible-graph" open>
              <summary>{text("Preuves détaillées par classeur et révision", "Evidence grouped by workbook revision")}</summary>
              {lineageGroups.map((group) => <section key={group.key} className="library-revision-group"><h4>{group.title || text("Révision inconnue", "Unknown revision")}</h4><ul>
                {group.nodes.map((node) => <li key={node.id}><strong>{lineageNodeLabels.get(node.id)}</strong>
                  <span>{text("Valeur", "Value")}: {node.value == null ? text("aucune valeur enregistrée", "no saved value") : searchText(node.value, english)} · {provenanceLabel(node.valueProvenance, english)}{node.valueProvenance === "saved_cache" && ` · ${text("aucune formule n’a été recalculée", "no formula was recalculated")}`}</span>
                  {node.formula && <span>{text("Formule", "Formula")}: <code>{node.formula}</code></span>}
                  {(node.labels || []).map((label, index) => <span key={`${node.id}:label:${index}`}>{roleLabel(label.role || label.method)} · {label.text} · {label.certainty === "heuristic" || !["table_column", "defined_name"].includes(label.method) ? text("inféré · certitude indicative", "inferred · indicative certainty") : text("déclaré", "declared")}{[label.sourceSheet || label.sheet, label.sourceCell || label.cell].filter(Boolean).length ? ` · ${text("source", "source")} ${[label.sourceSheet || label.sheet, label.sourceCell || label.cell].filter(Boolean).join("!")}` : ""}</span>)}
                </li>)}
              </ul></section>)}
              {lineageEdges.length > 0 && <ul>{lineageEdges.map((edge, index) => <li key={`${edge.source}:${edge.target}:${index}`}>{lineageNodeLabels.get(edge.source) || edge.source} → {lineageNodeLabels.get(edge.target) || edge.target} · {edge.kind === "external_link" ? text("lien externe", "external link") : edge.kind === "range" ? text("plage", "range") : text("dépendance de formule", "formula dependency")}{edge.resolutionStatus ? ` · ${statusLabel(edge.resolutionStatus)}` : ""}{edge.relationshipPath ? ` · ${text("référence", "reference")} ${edge.relationshipPath}` : ""}</li>)}</ul>}
              {lineageBoundaries.map((boundary, index) => {
                const nodeId = boundary.nodeId || boundary.node || boundary.sourceNode;
                return <p key={`boundary:${index}`}>{text("Frontière", "Boundary")} · {searchText(SEARCH_BOUNDARY_LABELS[boundary.reason] || boundary.reason, english)} · {lineageNodeLabels.get(nodeId) || ""}{boundary.stalePin ? ` · ${text("Épingle périmée", "Stale pin")}` : ""}{boundary.relationshipPath ? ` · ${boundary.relationshipPath}` : ""}{boundary.reference ? ` · ${boundary.reference}` : ""}</p>;
              })}
            </details>
          </div> : <p className="library-empty">{text("Choisissez un résultat pour suivre ses sources.", "Choose a result to follow its sources.")}</p>}
        </section>
      </div>
      <details className="library-links" open>
        <summary>{text("Liens externes", "External links")}</summary>
        <label>{text("État du lien", "Link status")}<select value={linkStatus} onChange={(event) => { setLinks([]); setLinkStatus(event.target.value); }}><option value="">{text("Tous", "All")}</option>{["ambiguous", "unresolved", "pinned", "resolved_by_name"].map((value) => <option key={value} value={value}>{statusLabel(value)}</option>)}</select></label>
        {linksCoverage?.complete === false && <p className="library-result-coverage is-incomplete">{text("Liste partielle :", "Partial list:")} {(linksCoverage.capsHit || []).map(capLabel).join(" · ")}</p>}
        {linksPending
          ? <p role="status">{text("Les liens apparaîtront lorsque l’indexation sera terminée.", "Links will appear when indexing completes.")}</p>
          : links.length === 0
            ? <p>{text("Aucun lien dans cette portée.", "No links in this scope.")}</p>
            : links.map((link) => {
          const pinned = link.status === "pinned" || link.pin;
          const targetOptions = revisions.filter((item) => item.revision?.sha256);
          const needsPin = ["ambiguous", "unresolved"].includes(link.status) || link.stalePin || pinned;
          const sourceRevision = link.sourceRevision || {};
          const sourceProject = revisions.find((item) => item.projectId === sourceRevision.projectId);
          const currentPin = link.pin && `${link.pin.projectId}:${link.pin.sha256}`;
          return <article key={link.id} className="library-link">
            <div><strong>{sourceProject?.filename || link.filename || sourceRevision.projectId || text("Classeur source", "Source workbook")}</strong>
              <span>{revisionLabel({ filename: sourceProject?.filename, projectId: sourceRevision.projectId, sha256: sourceRevision.sha256 })}</span>
              <span>{text("Cellule source", "Source cell")} · {link.sourceNode || text("non précisée", "not specified")}</span>
              <span>{link.reference || link.relationshipPath || text("Référence externe indisponible", "External reference unavailable")}</span>
              <span>{statusLabel(link.status)}{link.stalePin ? ` · ${text("épingle périmée", "stale pin")}` : ""}</span>
              {!!link.candidates?.length && <span>{text("Cibles possibles", "Possible targets")}: {link.candidates.map((candidate) => {
                const candidateProject = revisions.find((item) => item.projectId === candidate.projectId);
                return revisionLabel({ filename: candidateProject?.filename, projectId: candidate.projectId, sha256: candidate.sha256 });
              }).join(" · ")}</span>}
              {link.pin && <span>{text("Cible épinglée", "Pinned target")} · {revisionLabel({ filename: revisions.find((item) => item.projectId === link.pin.projectId)?.filename, projectId: link.pin.projectId, sha256: link.pin.sha256 })}</span>}
            </div>
            {needsPin && <>
              <label>{text("Révision cible", "Target revision")}<select aria-label={`${text("Révision cible pour", "Target revision for")} ${link.id}`} value={pinTargets[link.id] || currentPin || ""} onChange={(event) => setPinTargets((state) => ({ ...state, [link.id]: event.target.value }))}><option value="">{text("Choisir une révision", "Choose a revision")}</option>{targetOptions.map((item) => <option key={item.revision.id} value={`${item.projectId}:${item.revision.sha256}`}>{revisionLabel({ filename: item.filename, projectId: item.projectId, sha256: item.revision.sha256 })}</option>)}</select></label>
              <button type="button" onClick={() => pin(link)}>{pinned ? text("Confirmer cette cible", "Confirm this target") : text("Épingler cette cible", "Pin this target")}</button>
              {pinned && <button type="button" onClick={() => pin(link, true)}>{text("Retirer l’épingle", "Unpin")}</button>}
            </>}
          </article>;
        })}
        {linksPagination?.hasMore && <button type="button" onClick={() => refreshLinks(attachment?.scopeId, linkStatus, (linksPagination.offset || 0) + linksPagination.limit)}>{text("Afficher les liens suivants", "Show next links")}</button>}
      </details>
      <details id="library-trace" className="library-trace" open={intent === "ask" ? true : undefined}>
        <summary><h3 id="library-trace-heading" tabIndex={-1}>{text("Trace complète des opérations", "Full operation trace")}</h3></summary>
        {intent === "ask" && <AskTrace run={askRun} english={english} />}
        {intent === "search" && traceEvents.length === 0 && <p>{text("Aucune requête effectuée. La trace apparaîtra ici, y compris si aucun résultat n’est trouvé.", "No request yet. The trace will appear here, including when a search has no results.")}</p>}
        <ol>{traceEvents.map((event, index) => {
          const trace = event.trace || {};
          const scope = trace.scope || {};
          const counts = trace.counts || {};
          const operation = event.type === "search" ? text("Recherche", "Search") : event.type === "lineage" ? text("Lignage", "Lineage") : event.type === "occurrences" ? text("Occurrences", "Occurrences") : text("Import", "Import");
          const eventStatus = { pending: text("en cours", "pending"), completed: text("terminé", "completed"), failed: text("échec", "failed"), superseded: text("remplacé", "superseded") }[event.status];
          const directionName = event.detail === "upstream" ? text("En amont", "Upstream") : event.detail === "downstream" ? text("En aval", "Downstream") : event.detail;
          const revisionsInTrace = (trace.revisions || []).filter(Boolean).map((item) => revisionLabel({
            filename: revisions.find((revision) => revision.projectId === item.projectId)?.filename,
            projectId: item.projectId,
            sha256: item.sha256,
          }));
          return <li key={`${event.type}:${index}`}>
            <strong>{operation}{eventStatus ? ` · ${eventStatus}` : ""}</strong>
            <span>{event.type === "search"
              ? `${text("Requête", "Query")}: ${event.queryText} · ${text("Termes normalisés", "Normalized terms")}: ${(trace.terms || []).join(", ")} · ${text("Portée", "Scope")}: ${event.scopeLabel} · ${text("types", "types")}: ${(scope.kinds || []).map(kindLabel).join(", ")}`
              : event.type === "lineage"
                ? `${directionName} · ${text("révisions consultées", "revisions consulted")}: ${revisionsInTrace.join(", ") || text("en attente", "pending")}`
                : event.detail}</span>
            {event.error && <p className="library-trace-error">{text("Échec", "Failure")}: {event.error}</p>}
            {trace.normalization && <span>{text("Normalisation", "Normalization")}: {text("accents neutralisés, minuscules, tous les termes doivent correspondre", "accents folded, lower case, all terms must match")}</span>}
            {trace.scope?.scope && <span>{text("Périmètre", "Scope")}: {trace.scope.scope}</span>}
            {trace.limits && <span>{text("Budgets", "Budgets")}: {Object.entries(trace.limits).filter(([key]) => SEARCH_BUDGET_LABELS[key]).map(([key, value]) => `${searchText(SEARCH_BUDGET_LABELS[key], english)}: ${value}`).join(" · ")}</span>}
            <small>{trace.durationMs !== undefined ? traceTime(trace.durationMs) : ""}{trace.backend ? ` · ${trace.backend === "fts5" ? "FTS5" : text("recherche de secours", "fallback search")}` : ""}{trace.tool ? ` · ${operation}` : ""}{trace.capsHit?.length ? ` · ${text("arrêts", "stops")}: ${trace.capsHit.map(capLabel).join(", ")}` : ""}{trace.complete !== undefined ? ` · ${trace.complete ? text("couverture complète", "complete coverage") : text("couverture incomplète", "incomplete coverage")}` : event.complete !== undefined ? ` · ${event.complete ? text("couverture complète", "complete coverage") : text("couverture incomplète", "incomplete coverage")}` : ""}{trace.exhaustive !== undefined ? ` · ${trace.exhaustive ? text("parcours exhaustif", "exhaustive traversal") : text("parcours non exhaustif", "non-exhaustive traversal")}` : ""}</small>
            <details className="library-technical-trace"><summary>{text("Détails techniques", "Technical details")}</summary><pre>{JSON.stringify({ error: event.technicalError, generation: trace.generation, scope, revisions: trace.revisions, counts, limits: trace.limits, capsHit: trace.capsHit, direction: trace.direction }, null, 2)}</pre></details>
          </li>;
        })}</ol>
        {intent === "search" && <p className={`library-trace-last ${complete === false ? "is-incomplete" : ""}`}>{search ? (complete === false ? text("La recherche dans ce périmètre est incomplète ou tronquée.", "The scoped search is incomplete or truncated.") : text("La recherche du contenu indexé est complète dans ce périmètre.", "The indexed-content search is complete in this scope.")) : text("Aucune recherche exécutée.", "No search performed.")} {text("Une recherche complète du contenu indexé ne prouve pas l’absence d’une valeur dans les classeurs.", "A complete search of indexed content does not prove that a value is absent from workbooks.")}</p>}
      </details>
    </div>
  );
}

export function mountLibrarySearch(element, options = {}) {
  const root = createRoot(element);
  let update;
  function SearchRoot() {
    const [language, setLanguage] = useState(options.language || "fr");
    update = setLanguage;
    return <LibrarySearch apiBase={options.apiBase} language={language} onClose={options.onClose} />;
  }
  flushSync(() => root.render(<SearchRoot />));
  return {
    setLanguage: (language) => update(language),
    destroy: () => root.unmount(),
  };
}
