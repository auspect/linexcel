import React from "react";

const provenanceNames = {
  saved_input: ["valeur saisie et enregistrée", "saved input"],
  saved_cache: ["cache enregistré, sans recalcul", "saved cache, without recalculation"],
  saved_blank: ["cellule vide enregistrée", "saved blank cell"],
  source_range: ["plage du classeur", "workbook range"],
  unknown: ["origine non établie", "origin not established"],
};

const boundaryNames = {
  input_constant: ["Origine atteinte : donnée source", "Origin reached: source value"],
  unresolved_external: ["Classeur externe hors des sources consultables ou introuvable", "External workbook outside available sources or not found"],
  ambiguous_external: ["Plusieurs classeurs externes possibles", "Multiple possible external workbooks"],
  unsupported_reference: ["Référence non prise en charge", "Unsupported reference"],
  cache_only: ["Cache seul, source indisponible", "Cache only, source unavailable"],
  macro: ["Macro non exécutée", "Macro not executed"],
  power_query: ["Power Query non exécuté", "Power Query not executed"],
  cycle: ["Cycle détecté", "Cycle detected"],
  source_unavailable: ["Source indisponible", "Source unavailable"],
  source_index_incomplete: ["Index source incomplet", "Source index incomplete"],
  possible_unresolved_reader: ["Lecteur potentiel non résolu", "Possible unresolved reader"],
  depth: ["Limite de profondeur atteinte", "Depth limit reached"],
  node_budget: ["Limite de cellules atteinte", "Cell limit reached"],
  edge_budget: ["Limite de relations atteinte", "Relationship limit reached"],
  workbook_budget: ["Limite de classeurs atteinte", "Workbook limit reached"],
  time_budget: ["Limite de durée atteinte", "Time limit reached"],
  response_bytes: ["Limite de taille de réponse atteinte", "Response size limit reached"],
  link_budget: ["Limite de liens atteinte", "Link limit reached"],
  link_build_time: ["Durée de construction des liens atteinte", "Link build time reached"],
  index_rebuild_required: ["Réindexation requise", "Index rebuild required"],
};

// Excel shows at most 15 significant digits. Display only: receipts keep exact values.
export function displayValue(value) {
  return typeof value === "number" && Number.isFinite(value) && !Number.isInteger(value) ? String(Number(value.toPrecision(15))) : String(value);
}

// Stops that leave a proposal's lineage short of its sources. Reaching an input is an origin, not a limit.
export function lineageLimits(lineages, cells, english) {
  const labels = [];
  const add = label => { if (!labels.includes(label)) labels.push(label); };
  (cells || []).forEach((cell, index) => {
    const lineage = lineages?.[index];
    if (!lineage) { add(english ? "lineage unavailable" : "lignage indisponible"); return; }
    for (const boundary of lineage.boundaries || []) {
      if (boundary.reason === "input_constant") continue;
      add((boundaryNames[boundary.reason] || ["Autre arrêt du parcours", "Other traversal stop"])[english ? 1 : 0]);
    }
    for (const reason of lineage.coverage?.capsHit || []) add((boundaryNames[reason] || ["Limite de traitement", "Processing limit"])[english ? 1 : 0]);
    if (lineage.coverage?.complete !== true && !labels.length) add(english ? "coverage not established" : "couverture non établie");
  });
  return labels;
}

function valueText(value, english) {
  if (value == null) return english ? "no saved value" : "aucune valeur enregistrée";
  return typeof value === "object" ? JSON.stringify(value) : displayValue(value);
}

export function sourceProvenanceLabel(node, english) {
  if (node.formula && node.value == null) return english ? "formula cache unavailable, without recalculation" : "cache de formule indisponible, sans recalcul";
  return (provenanceNames[node.valueProvenance] || provenanceNames.unknown)[english ? 1 : 0];
}

function nodeLabel(node, english, withRevision = false) {
  if (!node) return english ? "cell not detailed in this traversal" : "cellule non détaillée dans ce parcours";
  const address = node.address || node.cell || "";
  const coordinate = address.includes("!") ? address : [node.sheet, address].filter(Boolean).join("!");
  return [node.filename, coordinate, withRevision && node.revision ? `${node.revision.projectId?.slice(0, 8) || ""}:${node.revision.sha256?.slice(0, 8) || ""}` : ""].filter(Boolean).join(" · ");
}

function SourceFacts({ node, english }) {
  const text = (fr, en) => english ? en : fr;
  const revision = node.revision || {};
  return <>
    <strong>{nodeLabel(node, english)}</strong>
    {(revision.projectId || revision.sha256) && <small>{text("Révision", "Revision")} · {revision.projectId?.slice(0, 8)} · {revision.sha256?.slice(0, 8)}</small>}
    {node.kind === "range" ? <span>{text("Plage source, sans valeur scalaire.", "Source range, without a scalar value.")}</span> : <span>{text("Valeur enregistrée", "Saved value")}: {valueText(node.value, english)} · {sourceProvenanceLabel(node, english)}</span>}
    {node.formula && <code>{node.formula}</code>}
    {(node.labels || []).map((label, index) => <span key={index}>{text("Contexte", "Context")}: {label.text} · {[label.sourceSheet || label.sheet, label.sourceCell || label.cell].filter(Boolean).join("!")} · {["table_column", "defined_name"].includes(label.method) ? text("déclaré dans le classeur", "declared in the workbook") : text("inféré, à vérifier", "inferred, to verify")}</span>)}
    {node.sourceValueProvenance && <details className="library-technical-trace"><summary>{text("Provenance de la valeur enregistrée", "Saved-value provenance")}</summary><pre>{JSON.stringify(node.sourceValueProvenance, null, 2)}</pre></details>}
  </>;
}

export function AnswerLineage({ root, lineage, english }) {
  const text = (fr, en) => english ? en : fr;
  const nodes = lineage?.nodes || [];
  const byId = new Map(nodes.map(node => [node.id, node]));
  const coverage = lineage?.coverage;
  const edges = lineage?.edges || [];
  const boundaries = lineage?.boundaries || [];
  const rootId = root.handle;
  return <details className="library-answer-proof" open data-proof-root={rootId}>
    <summary>{text("Lignage vérifié par linexcel", "Lineage verified by linexcel")}</summary>
    {!lineage ? <p>{text("Lignage indisponible pour cette proposition. Aucune couverture complète n’est établie.", "Lineage unavailable for this proposal. Complete coverage has not been established.")}</p> : <>
      <p className={`library-proof-status ${coverage?.complete === true ? "is-complete" : "is-incomplete"}`}>
        {coverage?.complete === true ? text("Lignage complet dans les sources autorisées.", "Lineage complete within the authorized sources.") : text("Lignage incomplet : consultez les origines et les arrêts ci-dessous.", "Lineage incomplete: check the origins and stops below.")}
        {" "}{coverage?.exhaustive === true ? text("Parcours exhaustif dans ce périmètre.", "Traversal exhaustive within this scope.") : text("Parcours non exhaustif.", "Traversal is not exhaustive.")}
      </p>
      <p>{text("Ces relations sont relues et calculées par linexcel pour cette cellule, indépendamment des recherches effectuées par le modèle.", "Linexcel rereads and computes these relationships for this cell, independently of the searches performed by the model.")}</p>
      <h5>{text("Relations de dépendance", "Dependency relationships")}</h5>
      <p className="help">{text("Sens des relations : source → cellule qui l’utilise.", "Relationship direction: source → cell that uses it.")}</p>
      {edges.length ? <ul className="library-proof-edges">{edges.map((edge, index) => <li key={`${edge.source}:${edge.target}:${index}`} data-source={edge.source} data-target={edge.target}>
        <span>{nodeLabel(byId.get(edge.source), english, true)}</span><span aria-hidden="true"> → </span><span className="library-proof-edge-direction">{text("alimente", "feeds")}</span><span>{nodeLabel(byId.get(edge.target), english, true)}</span>
        {edge.kind === "external_link" && <small>{text("Lien externe", "External link")}{edge.relationshipPath ? ` · ${edge.relationshipPath}` : ""}{edge.resolutionStatus === "resolved_by_name" ? ` · ${text("correspondance par nom, révision d’origine non établie", "matched by name, original revision not established")}` : ""}</small>}
      </li>)}</ul> : <p>{text("Aucune relation supplémentaire renvoyée pour cette cellule.", "No additional relationship returned for this cell.")}</p>}
      <h5>{text("Cellules et provenance", "Cells and provenance")}</h5>
      <ul className="library-proof-nodes">{nodes.map(node => <li key={node.id} data-lineage-node={node.id}><SourceFacts node={node} english={english} /></li>)}</ul>
      <h5>{text("Origines et arrêts du parcours", "Traversal origins and stops")}</h5>
      {boundaries.length ? <ul className="library-proof-boundaries">{boundaries.map((boundary, index) => <li key={index}>
        <strong>{(boundaryNames[boundary.reason] || ["Autre arrêt du parcours", "Other traversal stop"])[english ? 1 : 0]}</strong>
        <span>{nodeLabel(byId.get(boundary.node || boundary.nodeId || boundary.sourceNode), english, true)}</span>
        {boundary.reference && <code>{boundary.reference}</code>}
        {boundary.relationshipPath && <span>{boundary.relationshipPath}</span>}
        {boundary.stalePin && <span>{text("La cible épinglée n’est plus à jour.", "The pinned target is no longer current.")}</span>}
      </li>)}</ul> : <p>{text("Aucune origine ou frontière supplémentaire signalée.", "No additional origin or boundary reported.")}</p>}
      {!!coverage?.capsHit?.length && <p>{text("Limites atteintes", "Limits reached")}: {coverage.capsHit.map(reason => (boundaryNames[reason] || ["Limite de traitement", "Processing limit"])[english ? 1 : 0]).join(" · ")}</p>}
    </>}
  </details>;
}

export function CalculationReceipt({ receipt, english }) {
  const text = (fr, en) => english ? en : fr;
  if (!receipt) return <p>{text("Recette reproductible indisponible pour ce calcul.", "Reproducible recipe unavailable for this calculation.")}</p>;
  return <details className="library-calculation-receipt" open>
    <summary>{text("Code et sources du calcul", "Calculation code and sources")}</summary>
    <p>{text("Ce code est produit par linexcel à partir de l’opération et des valeurs figées ci-dessous. Il est affiché sans être exécuté dans le navigateur.", "Linexcel produces this code from the operation and saved snapshot below. It is displayed without executing it in the browser.")}</p>
    <p>{text("Moteur", "Engine")}: <code>{receipt.engine}</code> · {text("version", "version")} {receipt.engineVersion}</p>
    <pre className="library-receipt-code"><code>{receipt.code}</code></pre>
    <p>SHA-256 : <code>{receipt.codeSha256}</code></p>
    <h5>{text("Sources figées du calcul", "Frozen calculation sources")}</h5>
    <ul className="library-proof-nodes">{(receipt.operands || []).map((node, index) => <li key={`${node.handle}:${index}`}><SourceFacts node={node} english={english} /></li>)}</ul>
  </details>;
}
