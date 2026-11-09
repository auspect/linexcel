// Aggregate bounded relationships without inventing dependency edges.
function aggregateEdges(edges, representative, distinguishKinds = false) {
  const links = new Map();
  for (const edge of edges) {
    const source = representative(edge.source), target = representative(edge.target);
    if (source === target && distinguishKinds && source !== edge.source) continue;
    const key = JSON.stringify([source, target, distinguishKinds ? edge.kind : null]);
    if (!links.has(key)) links.set(key, { ...edge, source, target, weight: 0 });
    links.get(key).weight += edge.weight || 1;
  }
  return [...links.values()];
}

// A collapsed revision represents its actual members; expanding one revision
// retains other workbooks as context. Internal relationships are counted on
// the collapsed card, never drawn as fabricated membership dependencies.
export function revisionLineageView(lineage, expandedRevision = null) {
  const groups = new Map(), membership = new Map(), boundaries = new Map();
  for (const boundary of lineage.boundaries || []) {
    const id = boundary.nodeId || boundary.node || boundary.sourceNode;
    if (!boundaries.has(id)) boundaries.set(id, []);
    boundaries.get(id).push(boundary);
  }
  for (const original of lineage.nodes || []) {
    const revision = original.revision || original.workbookRevision || {};
    const id = `revision:${revision.id || `${original.projectId}:${revision.sha256 || original.sha256}`}`;
    if (!groups.has(id)) groups.set(id, {
      id, revision, filename: original.filename || original.projectName,
      members: [], internalEdges: 0, boundaries: [],
    });
    const node = { ...original, boundaries: boundaries.get(original.id) || [], revisionGroupId: id };
    groups.get(id).members.push(node);
    groups.get(id).boundaries.push(...node.boundaries);
    membership.set(node.id, id);
  }
  for (const edge of lineage.edges || []) {
    if (membership.get(edge.source) === membership.get(edge.target)) {
      const group = groups.get(membership.get(edge.source));
      if (group) group.internalEdges += edge.weight || 1;
    }
  }
  const expanded = groups.has(expandedRevision) ? expandedRevision : null;
  const nodes = [...groups.values()].flatMap(group => group.id === expanded
    ? group.members
    : [{ ...group, kind: "revision_group", collapsed: true }]);
  const representative = id => membership.get(id) === expanded ? id : membership.get(id);
  const sourceEdges = (lineage.edges || []).filter(e => membership.has(e.source) && membership.has(e.target));
  const edges = aggregateEdges(sourceEdges, representative, true);
  return {
    nodes, edges, groups: [...groups.values()], expandedRevision: expanded,
    internalEdgesGrouped: [...groups.values()].filter(g => g.id !== expanded).reduce((sum, g) => sum + g.internalEdges, 0),
  };
}

// Work only on the bounded response. First summarize independent copied-formula
// siblings, then their equal-adjacency inputs. Never fold a dependency chain.
export function compressNeighborhood(nodes, edges, centerId) {
  let grouped = 0;
  for (let pass = 0; pass < 2; pass++) {
  const adjacency = new Map(nodes.map(n => [n.id, []]));
  const outgoing = new Map(nodes.map(n => [n.id, []]));
  for (const e of edges) {
    adjacency.get(e.source)?.push(`out:${e.target}`);
    adjacency.get(e.target)?.push(`in:${e.source}`);
    outgoing.get(e.source)?.push(e.target);
  }
  const buckets = new Map();
  for (const n of nodes) {
    const cell = /^([A-Z]+)(\d+)$/.exec(n.cell || '');
    if (n.id === centerId || !cell || !['cell','formula','input'].includes(n.kind)) continue;
    const key = JSON.stringify([n.sheet, cell[1], n.patternId || n.kind,
      pass === 0 && n.patternId ? [] : adjacency.get(n.id).slice().sort()]);
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(n);
  }
  const replacements = new Map(), groups = [];
  for (const members of buckets.values()) {
    if (members.length < 4) continue;
    // The same R1C1 pattern can describe a running total. Reachability, not
    // just a matching formula, distinguishes independent siblings from it.
    const memberIds = new Set(members.map(n => n.id));
    const dependent = members.some(n => {
      const pending = [...(outgoing.get(n.id) || [])], visited = new Set();
      while (pending.length) {
        const id = pending.pop();
        if (memberIds.has(id)) return true;
        if (visited.has(id)) continue;
        visited.add(id); pending.push(...(outgoing.get(id) || []));
      }
      return false;
    });
    if (dependent) continue;
    members.sort((a,b) => Number(/\d+$/.exec(a.cell)[0])-Number(/\d+$/.exec(b.cell)[0]));
    const first = members[0], last = members.at(-1);
    const start = Number(/\d+$/.exec(first.cell)[0]), end = Number(/\d+$/.exec(last.cell)[0]);
    const id = `aggregate:${first.id}`;
    groups.push({id, aggregate:true, kind:'cell_group', sheet:first.sheet,
      label:end-start+1===members.length ? `${first.sheet}!${first.cell}:${last.cell}`
        : `${first.sheet} · ${first.cell.replace(/\d+$/,'')}`,
      members:members.map(n=>n.id), anchorId:first.id, patternId:first.patternId,
      patternTotal:Math.max(...members.map(n=>n.patternMemberCount||0))});
    members.forEach(n=>replacements.set(n.id,id));
  }
  if (!groups.length) continue;
  grouped += replacements.size;
  const links = aggregateEdges(edges, id => replacements.get(id) || id);
  nodes = [...nodes.filter(n=>!replacements.has(n.id)),...groups];
  edges = links;
  }
  return {nodes, edges, grouped};
}
