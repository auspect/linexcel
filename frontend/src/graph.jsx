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
        className: e.membership ? "membership" : "",
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
