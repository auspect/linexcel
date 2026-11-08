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
  const links = new Map();
  for (const e of edges) {
    const source=replacements.get(e.source)||e.source, target=replacements.get(e.target)||e.target;
    const key=JSON.stringify([source,target]);
    if (!links.has(key)) links.set(key,{...e,source,target,weight:0});
    links.get(key).weight += e.weight||1;
  }
  nodes = [...nodes.filter(n=>!replacements.has(n.id)),...groups];
  edges = [...links.values()];
  }
  return {nodes, edges, grouped};
}
