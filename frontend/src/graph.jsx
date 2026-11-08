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
  getSmoothStepPath,
} from "@xyflow/react";
import dagre from "@dagrejs/dagre";
import "@xyflow/react/dist/style.css";
import "./graph.css";
import "./shell.css";
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
  points[0] = { x: sourceX, y: sourceY };
  points[points.length - 1] = { x: targetX, y: targetY };
  const packed=data?.packed?getSmoothStepPath({sourceX,sourceY,targetX,targetY,sourcePosition,targetPosition}):null;
  const path = packed?packed[0]:points.map((p, i) => `${i ? "L" : "M"}${p.x},${p.y}`).join(" "),
    center = points[Math.floor(points.length / 2)];
  return (
    <>
      <BaseEdge
        id={id}
        path={path}
        markerEnd={markerEnd}
        label={label}
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

function layered(nodes, edges, portrait, aspect) {
  const owned=nodes.filter(n=>n.data.tone!=='outside'),outside=nodes.find(n=>n.data.tone==='outside');
  const ownedIds=new Set(owned.map(n=>n.id));
  const leaf=owned.length>1&&owned.every(n=>n.data.original.nodeId)&&!edges.some(e=>ownedIds.has(e.source)&&ownedIds.has(e.target));
  if(leaf){
    const columns=Math.max(1,Math.min(6,Math.ceil(Math.sqrt(owned.length*aspect/2.23))));
    owned.forEach((n,i)=>{n.position={x:(i%columns)*272,y:Math.floor(i/columns)*144};});
    const width=columns*272-40,height=Math.ceil(owned.length/columns)*144-40;
    if(outside)outside.position={x:width+110,y:Math.max(0,(height-104)/2)};
    const lookup=new Map(nodes.map(n=>[n.id,n]));
    for(const edge of edges){
      const a=lookup.get(edge.source),b=lookup.get(edge.target),s={x:a.position.x+232,y:a.position.y+52},t={x:b.position.x,y:b.position.y+52};
      let points;
      if(a===outside){const busY=b.position.y-18;points=[s,{x:s.x+24,y:s.y},{x:s.x+24,y:a.position.y-20},{x:width+42,y:a.position.y-20},{x:width+42,y:busY},{x:b.position.x-18,y:busY},{x:b.position.x-18,y:t.y},t];}
      else points=[s,{x:s.x+18,y:s.y},{x:s.x+18,y:a.position.y+122},{x:width+42,y:a.position.y+122},{x:width+42,y:t.y},t];
      edge.data={...edge.data,points};edge.sourceHandle='source-right';edge.targetHandle='target-left';
    }
    return nodes;
  }
  const graph = new dagre.graphlib.Graph({ multigraph: true });
  graph.setGraph({
    rankdir: portrait ? "TB" : "LR",
    ranksep: 100,
    nodesep: 35,
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
  // Pack unusually broad ranks after Dagre assigns their dependency layers.
  // The bounded overview remains useful even when siblings cannot be grouped.
  const ranks=new Map();
  for(const n of nodes){const p=graph.node(n.id),key=portrait?p.y:p.x;
    if(!ranks.has(key))ranks.set(key,[]);ranks.get(key).push(n);}
  if([...ranks.values()].some(rank=>rank.length>8)){
    let offset=30;
    for(const [,rank] of [...ranks].sort((a,b)=>a[0]-b[0])){
      const columns=Math.min(6,Math.ceil(Math.sqrt(rank.length*aspect/2.23)));
      const rows=Math.ceil(rank.length/columns);
      rank.forEach((n,i)=>{const p=graph.node(n.id);
        p.x=(portrait?30:offset)+(i%columns)*272+116;
        p.y=(portrait?offset:30)+Math.floor(i/columns)*144+52;});
      offset+=(portrait?rows*144:columns*272)+100;
    }
    for(const edge of edges){const a=graph.node(edge.source),b=graph.node(edge.target);
      const horizontal=Math.abs(b.x-a.x)>=Math.abs(b.y-a.y);
      const side=horizontal?(a.x<b.x?'right':'left'):(a.y<b.y?'bottom':'top');
      const reverse={left:'right',right:'left',top:'bottom',bottom:'top'}[side];
      edge.sourceHandle='source-'+side;edge.targetHandle='target-'+reverse;
      edge.data={...edge.data,packed:true};
    }
    return nodes.map(n=>({...n,position:{x:graph.node(n.id).x-116,y:graph.node(n.id).y-52}}));
  }
  const side=(p,center)=>Math.abs((p.x-center.x)/116)>=Math.abs((p.y-center.y)/52)?p.x<center.x?'left':'right':p.y<center.y?'top':'bottom';
  edges.forEach((edge) => {
    const points=graph.edge({v:edge.source,w:edge.target,name:edge.id}).points;
    edge.sourceHandle='source-'+side(points[0],graph.node(edge.source));
    edge.targetHandle='target-'+side(points[points.length-1],graph.node(edge.target));
    edge.data = {
      ...edge.data,
      points,
    };
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
      const placed = layered(mapped, links, aspect<.85,aspect);
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
