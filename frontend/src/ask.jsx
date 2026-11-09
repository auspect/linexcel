import React, { useEffect, useRef, useState } from "react";

const failures = {
  tool_calls_required: ["Ce point de terminaison ne fournit pas les appels d’outils requis. Choisissez un modèle compatible avec les outils.", "This endpoint does not provide the required tool calls. Choose a tool-capable model."],
  truncated_or_unsupported_calls: ["Les appels du modèle sont tronqués ou non pris en charge.", "The model calls were truncated or unsupported."],
  invalid_or_unavailable_tool_provider: ["Le modèle est indisponible ou a envoyé un appel d’outil invalide.", "The model is unavailable or sent an invalid tool call."],
  scope_or_source_changed: ["La portée ou une révision a changé. Relancez la question.", "The scope or a revision changed. Ask again."],
  cancelled: ["Question annulée.", "Question cancelled."],
  interrupted: ["Question interrompue au redémarrage.", "Question interrupted by a restart."],
};
const toolNames = {
  list_workbooks: ["Classeurs autorisés", "Authorized workbooks"],
  list_revisions: ["Révisions autorisées", "Authorized revisions"],
  search_catalog: ["Recherche du catalogue", "Catalog search"],
  describe_sheet: ["Repérage de la feuille", "Sheet inspection"],
  read_range: ["Lecture d’une petite plage", "Small range read"],
  get_cell: ["Vérification d’une cellule", "Cell verification"],
  find_value: ["Recherche d’une valeur", "Value search"],
  trace_lineage: ["Parcours du lignage", "Lineage traversal"],
  list_external_links: ["Liens externes", "External links"],
  compute: ["Calcul par linexcel", "Calculation by linexcel"],
  submit_answer: ["Vérification de la proposition", "Proposal verification"],
  model: ["Sélection du modèle", "Model selection"],
  failure: ["Échec de l’opération", "Operation failed"],
  finish: ["Fin de la recherche", "Search finished"],
};
const provenanceNames = {
  saved_input: ["valeur saisie et enregistrée", "saved input"],
  saved_cache: ["cache enregistré, sans recalcul", "saved cache, without recalculation"],
  saved_blank: ["cellule vide enregistrée", "saved blank cell"],
  unknown: ["origine non établie", "origin not established"],
};
const labelNames = {
  left_text: ["texte à gauche, contexte inféré", "text on the left, inferred context"],
  above_text: ["texte au-dessus, contexte inféré", "text above, inferred context"],
  header_band: ["en-tête inféré, heuristique", "inferred header, heuristic"],
  table_column: ["colonne déclarée dans le classeur", "column declared in the workbook"],
  defined_name: ["nom déclaré dans le classeur", "name declared in the workbook"],
};
const operationNames = {
  sum: ["Somme", "Sum"], difference: ["Différence", "Difference"], ratio: ["Rapport", "Ratio"],
  min: ["Minimum", "Minimum"], max: ["Maximum", "Maximum"], mean: ["Moyenne", "Mean"], count: ["Nombre de valeurs numériques", "Numeric value count"],
};

export function AskTrace({ run, english }) {
  const text = (fr, en) => english ? en : fr;
  if (!run) return <p>{text("Aucune question exécutée.", "No question run yet.")}</p>;
  const trace = run.trace;
  const reason = trace?.reason;
  return <div className="library-ask-trace" aria-live="polite">
    {failures[reason] && <p role="status">{failures[reason][english ? 1 : 0]}</p>}
    <ol>{(trace?.steps || []).map((step, index) => <li key={index}>
      <strong>{(toolNames[step.tool] || [step.tool, step.tool])[english ? 1 : 0]} · {step.verdict === "failed" ? text("échec", "failed") : step.tool === "finish" ? text("terminée", "finished") : text("effectué", "completed")}</strong>
      {step.durationMs !== undefined && <small>{Math.round(step.durationMs)} ms{step.tokens ? ` · ${step.tokens.input_tokens + step.tokens.output_tokens} tokens` : ""}</small>}
      {step.summary?.counts && <span>{Object.entries(step.summary.counts).map(([key, count]) => `${count} ${{ cells: text("cellules", "cells"), results: text("résultats", "results"), nodes: text("nœuds", "nodes"), edges: text("relations", "relationships"), revisions: text("révisions", "revisions"), links: text("liens", "links") }[key] || key}`).join(" · ")}</span>}
      {step.tool === "finish" && <p className="library-ask-last">{step.summary?.searchCompleted
        ? text("La recherche délimitée a terminé ses opérations dans cette portée.", "The bounded search completed its operations in this scope.")
        : text("La recherche dans cette portée est incomplète ou tronquée.", "The scoped search is incomplete or truncated.")} {text("Cela ne prouve pas l’absence d’une réponse dans les classeurs.", "This does not prove that an answer is absent from the workbooks.")}</p>}
      <details className="library-technical-trace"><summary>{text("Détails techniques", "Technical details")}</summary><pre>{JSON.stringify({ tool: step.tool, arguments: step.arguments, summary: step.summary, budgetUse: step.budgetUse, verdict: step.verdict }, null, 2)}</pre></details>
    </li>)}</ol>
    {trace?.budgetUse && <p>{trace.budgetUse.steps} {text("étapes", "steps")} · {trace.budgetUse.tokens} tokens · {trace.budgetUse.resultBytes} {text("octets de résultats", "result bytes")}</p>}
    {trace?.scope?.revisions && <details><summary>{text("Révisions autorisées pour cette question", "Revisions authorized for this question")}</summary><ul>{trace.scope.revisions.map(rev => <li key={rev.id}>{rev.filename} · {rev.projectId.slice(0, 8)} · {rev.sha256.slice(0, 8)}</li>)}</ul></details>}
  </div>;
}

export function AskQuestion({ request, scope, english, onRun, onEvidence, onClear }) {
  const text = (fr, en) => english ? en : fr;
  const [config, setConfig] = useState(null);
  const [question, setQuestion] = useState("");
  const [disclosure, setDisclosure] = useState(null);
  const [run, setRun] = useState(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [feedback, setFeedback] = useState(null);
  const identity = useRef(0), activeTask = useRef(null), timer = useRef(null), mounted = useRef(true);
  const updateRun = value => { setRun(value); onRun(value); };
  useEffect(() => {
    request("library/ask/config").then(value => mounted.current && setConfig(value)).catch(() => mounted.current && setError(text("Configuration indisponible.", "Configuration unavailable.")));
    return () => {
      mounted.current = false;
      identity.current += 1;
      clearTimeout(timer.current);
      if (activeTask.current) request(`tasks/${activeTask.current}/cancel`, { method: "POST", body: "{}" }).catch(() => {});
    };
  }, []);
  const clear = () => {
    identity.current += 1;
    clearTimeout(timer.current);
    if (activeTask.current) request(`tasks/${activeTask.current}/cancel`, { method: "POST", body: "{}" }).catch(() => {});
    activeTask.current = null;
    setBusy(false); setDisclosure(null); updateRun(null); setFeedback(null); setError(""); onClear();
  };
  const accept = (value, id) => {
    if (!mounted.current || id !== identity.current) return;
    runId.current = value.task.id;
    updateRun(value.run);
    if (["succeeded", "failed", "cancelled"].includes(value.task.status)) {
      setBusy(false); activeTask.current = null;
      if (value.run.status === "answer_proposed") onEvidence(value.run.answer);
    } else timer.current = setTimeout(() => poll(value.task.id, id), 500);
  };
  const poll = async (taskId, id) => {
    try { accept(await request(`library/ask/${taskId}`), id); }
    catch { if (mounted.current && id === identity.current) { setError(text("Impossible de lire cette exécution. Relancez la question.", "Cannot read this run. Ask again.")); setBusy(false); } }
  };
  const start = async (prepared, consent, id) => {
    const value = await request("library/ask", { method: "POST", body: JSON.stringify({ token: prepared.token, remoteConsent: consent }) });
    if (!mounted.current || id !== identity.current) {
      request(`tasks/${value.task.id}/cancel`, { method: "POST", body: "{}" }).catch(() => {}); return;
    }
    activeTask.current = value.task.id;
    setDisclosure(null); setBusy(true);
    await poll(value.task.id, id);
  };
  const prepare = async event => {
    event.preventDefault(); clear(); setBusy(true);
    const id = identity.current;
    try {
      const prepared = await request("library/ask/prepare", { method: "POST", body: JSON.stringify({ question, ...scope }) });
      if (!mounted.current || id !== identity.current) return;
      if (prepared.local) await start(prepared, false, id);
      else { setDisclosure(prepared); setBusy(false); }
    } catch (cause) { if (mounted.current && id === identity.current) { setError(cause.message); setBusy(false); } }
  };
  const cancel = async () => {
    if (!activeTask.current) { clear(); return; }
    try { await request(`tasks/${activeTask.current}/cancel`, { method: "POST", body: "{}" }); }
    catch { setError(text("Annulation indisponible. Réessayez.", "Cancellation unavailable. Try again.")); }
  };
  const decide = async decision => {
    try {
      const value = await request(`library/ask/${runId.current}/feedback`, { method: "POST", body: JSON.stringify({ decision }) });
      setFeedback(value.feedback.at(-1).decision);
    } catch (cause) { setError(cause.message); }
  };
  const runId = useRef(null);
  const answer = run?.status === "answer_proposed" ? run.answer : null;
  return <section className="library-ask-state" aria-label={text("Poser une question", "Ask a question")}>
    <p className="help">{text("Le modèle cherche et propose des cellules. Les valeurs et le lignage sont relus par linexcel ; leur pertinence reste à confirmer.", "The model searches and proposes cells. Linexcel rereads values and lineage; relevance still needs confirmation.")}</p>
    {config?.endpoint && <p className="library-ask-endpoint">{config.local ? text("Modèle local", "Local model") : text("Modèle distant, autorisation par question", "Remote model, consent per question")} · {config.model} · {config.endpoint}</p>}
    {config && !config.available && <p role="status">{text("Aucun modèle configuré pour les questions. La recherche reste disponible.", "No model configured for questions. Search remains available.")}</p>}
    <form onSubmit={prepare}>
      <label htmlFor="library-ask-question">{text("Votre question", "Your question")}</label>
      <textarea id="library-ask-question" rows={2} maxLength={2000} value={question} disabled={!config?.available} onChange={event => { clear(); setQuestion(event.target.value); }} />
      <div className="library-ask-actions"><button type="submit" className={disclosure ? "" : "primary"} disabled={!config?.available || busy || !question.trim() || Boolean(disclosure)}>{text("Chercher une réponse", "Search for an answer")}</button>{busy && <button type="button" onClick={cancel}>{text("Annuler", "Cancel")}</button>}</div>
    </form>
    {busy && <p role="status">{text("Recherche en cours", "Searching")} · {run?.trace?.budgetUse?.steps || 0} {text("étapes effectuées. La trace est disponible ci-dessus.", "steps completed. The trace is accessible above.")}</p>}
    {error && <p role="alert">{error}</p>}
    {disclosure && <section className="library-ask-consent" aria-label={text("Autorisation distante", "Remote consent")}>
      <h3>{text("Autoriser cette question à distance", "Allow this question remotely")}</h3>
      <p>{text("La question, les valeurs, les formules, les libellés et les noms des feuilles et classeurs consultés quitteront cette machine vers", "The question, values, formulas, labels and names of consulted sheets and workbooks will leave this machine for")} <strong>{disclosure.endpoint}</strong> · {disclosure.model}.</p>
      <p>{disclosure.question}</p><ul>{disclosure.revisions.map(rev => <li key={rev.id}>{rev.filename} · {rev.projectId.slice(0, 8)} · {rev.sha256.slice(0, 8)}</li>)}</ul>
      <button type="button" className="primary" onClick={async () => { setBusy(true); try { await start(disclosure, true, identity.current); } catch(cause) { setError(cause.message); setBusy(false); } }}>{text("Autoriser pour cette question", "Allow for this question")}</button>
      <button type="button" onClick={() => setDisclosure(null)}>{text("Revenir sans envoyer", "Return without sending")}</button>
    </section>}
    {answer && <section className="library-ask-answer" aria-label={text("Réponse proposée", "Proposed answer")}>
      <h3>{text("Réponse proposée", "Proposed answer")}</h3>
      <p><strong>{text("Valeurs, formules et lignage vérifiés dans les index.", "Values, formulas and lineage verified in the indexes.")}</strong> {text("Pertinence à confirmer. Les libellés inférés restent heuristiques ; les limites du lignage sont indiquées ci-dessous.", "Relevance to confirm. Inferred labels remain heuristic; lineage limits are shown below.")}</p>
      {answer.cells.map((cell, index) => <article className="library-ask-cell" key={cell.handle}>
        <strong>{cell.filename} · {cell.sheet}!{cell.address}</strong><small>{cell.revision.projectId.slice(0, 8)} · {cell.revision.sha256.slice(0, 8)}</small>
        <span>{text("Valeur enregistrée", "Saved value")}: {cell.value == null ? text("aucune valeur enregistrée", "no saved value") : String(cell.value)}</span>
        <span>{text("Provenance", "Provenance")}: {(provenanceNames[cell.valueProvenance] || provenanceNames.unknown)[english ? 1 : 0]}</span>
        {cell.formula && <code>{cell.formula}</code>}
        {(cell.labels || []).map((label, i) => <span key={i}>{label.text} · {label.sourceSheet || label.sheet}!{label.sourceCell || label.cell} · {(labelNames[label.method] || ["contexte inféré, à vérifier", "inferred context, to verify"])[english ? 1 : 0]}</span>)}
        <button type="button" onClick={() => onEvidence(answer, index)}>{text("Voir son lignage et ses limites", "View its lineage and limits")}</button>
      </article>)}
      {answer.computation && <article className="library-calculation" data-node-id={answer.computation.id}><h4>{text("Calculé par linexcel pour cette réponse, présent dans aucun classeur", "Calculated by linexcel for this answer, present in no workbook")}</h4><p>{operationNames[answer.computation.op]?.[english ? 1 : 0]} · {answer.computation.error ? text("Calcul indisponible : vérifiez les valeurs et erreurs des opérandes.", "Calculation unavailable: check operand values and errors.") : String(answer.computation.value)}</p><p>{text("Opérandes", "Operands")}: {answer.computation.operands.map(cell => `${cell.filename} · ${cell.sheet}!${cell.address}`).join(" ; ")}</p><p>{text("Un calcul valide ne prouve pas qu’il répond à la question.", "Valid arithmetic does not prove that it answers the question.")}</p></article>}
      {answer.rationale && <details><summary>{text("Le raisonnement du modèle, non vérifié", "The model’s reasoning, unverified")}</summary><p>{answer.rationale}</p></details>}
      <div className="library-ask-actions"><button type="button" onClick={() => decide("confirm")}>{text("Confirmer la pertinence", "Confirm relevance")}</button><button type="button" onClick={() => decide("reject")}>{text("Rejeter la proposition", "Reject proposal")}</button></div>
      {feedback && <p role="status">{feedback === "confirm" ? text("Pertinence confirmée par vous.", "Relevance confirmed by you.") : text("Proposition rejetée par vous.", "Proposal rejected by you.")}</p>}
    </section>}
  </section>;
}
