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

const Card = memo(({ data, selected }) => (
  <article
    data-no-i18n=""
    className={`flow-card ${data.tone || ""} ${selected ? "is-selected" : ""}`}
    title={data.title}
  >
    <Handle
      type="target"
      position={data.portrait ? Position.Top : Position.Left}
      isConnectable={false}
    />
    <span className="flow-kind">{data.kindLabel}</span>
    <strong>{data.title}</strong>
    <span className="flow-subtitle">{data.subtitle}</span>
    <Handle
      type="source"
      position={data.portrait ? Position.Bottom : Position.Right}
      isConnectable={false}
    />
  </article>
));
const nodeTypes = { cell: Card, formulaGroup: Card, sheet: Card };
function LayeredEdge({
  id,
  sourceX,
  sourceY,
  targetX,
  targetY,
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
  const path = points.map((p, i) => `${i ? "L" : "M"}${p.x},${p.y}`).join(" "),
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

function layered(nodes, edges, portrait) {
  const graph = new dagre.graphlib.Graph({ multigraph: true });
  graph.setGraph({
    rankdir: portrait ? "TB" : "LR",
    ranksep: 100,
    nodesep: 35,
    marginx: 30,
    marginy: 30,
  });
  graph.setDefaultEdgeLabel(() => ({}));
  nodes.forEach((node) => graph.setNode(node.id, { width: 232, height: 104 }));
  edges.forEach((edge) =>
    graph.setEdge(
      edge.source,
      edge.target,
      { weight: Math.min(Number(edge.label) || 1, 100) },
      edge.id,
    ),
  );
  dagre.layout(graph);
  edges.forEach((edge) => {
    edge.data = {
      ...edge.data,
      points: graph.edge({ v: edge.source, w: edge.target, name: edge.id })
        .points,
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
      const observer = new ResizeObserver(([entry]) =>
        element.style.setProperty(
          "--graph-hud-height",
          `${entry.contentRect.height}px`,
        ),
      );
      if (hudRef.current) observer.observe(hudRef.current);
      return () => observer.disconnect();
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
            nodes: applyNodeChanges(changes, current.nodes),
          };
          setModel(current);
        }}
        onNodeClick={(_, node) => options.onSelect(node.data.original)}
        onMove={(_, viewport) => options.onMove?.(viewport)}
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
        <Controls showInteractive={false} fitViewOptions={api.fitOptions()} />
        <Background gap={22} size={1} color="#ccd9d4" />
        <Panel ref={hudRef} position="top-left" className="graph-hud">
          <Dock elements={options.toolbar} />
        </Panel>
        <Panel position="top-left" className="graph-breadcrumb-panel">
          <Dock elements={options.breadcrumbs} />
        </Panel>
        <Panel position="bottom-left" className="graph-status-panel">
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
      const links = edges.map((e, i) => ({
        ...e,
        id: e.id || `edge-${i}`,
        type: "layered",
        markerEnd: { type: MarkerType.ArrowClosed },
        label: e.label ?? (e.weight ? String(e.weight) : undefined),
        className: e.membership ? "membership" : "",
      }));
      const placed = layered(mapped, links, element.clientWidth < 550);
      if (!fit)
        for (const node of placed) {
          const previous = current.nodes.find((n) => n.id === node.id);
          if (previous) node.position = previous.position;
        }
      current = { nodes: placed, edges: links, language, fit, focus };
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
    setViewport: (viewport) => instance?.setViewport(viewport),
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
      const trail =
        breadcrumbs && !breadcrumbs.hidden
          ? breadcrumbs.getBoundingClientRect().height
          : 0;
      return {
        padding: {
          top: `${Math.min(hud + trail + 26, element.clientHeight * 0.48)}px`,
          right: "24px",
          bottom: "34px",
          left: "24px",
        },
        minZoom: 0.001,
        maxZoom: 1.2,
        duration: 0,
      };
    },
    fit: () => instance?.fitView(api.fitOptions()),
    focus: (id) => {
      const n = current.nodes.find((n) => n.id === id || n.data.nodeId === id);
      if (n)
        instance?.setCenter(n.position.x + 116, n.position.y + 52, {
          zoom: 1,
          duration: 0,
        });
    },
    select: (id) => {
      const n = api.getNode(id);
      if (n) options.onSelect(n.data.original);
    },
    zoom: (factor) => instance?.zoomTo((instance?.getZoom() || 1) * factor),
    panBy: (delta) => {
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
