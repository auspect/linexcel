import React, { useEffect, useRef, useState } from "react";
import { AnswerLineage, CalculationReceipt, displayValue, lineageLimits, sourceProvenanceLabel } from "./ask-evidence.jsx";

const failures = {
  endpoint_redirect_refused: ["La redirection du fournisseur a été refusée. Aucun contenu n’a été transmis à cette destination.", "The provider redirect was refused. No content was forwarded to that destination."],
  approved_configuration_changed: ["Le fournisseur ou le modèle a changé depuis l’autorisation. Préparez à nouveau la question.", "The endpoint or model changed after approval. Prepare the question again."],
  lineage_state_changed: ["Les liens ou le catalogue ont changé depuis cette proposition. Relancez la question pour vérifier son lignage actuel.", "Links or the catalog changed after this proposal. Ask again to verify its current lineage."],
  tool_calls_required: ["Ce point de terminaison ne fournit pas les appels d’outils requis. Choisissez un modèle compatible avec les outils.", "This endpoint does not provide the required tool calls. Choose a tool-capable model."],
  truncated_or_unsupported_calls: ["Les appels du modèle sont tronqués ou non pris en charge.", "The model calls were truncated or unsupported."],
  invalid_or_unavailable_tool_provider: ["Le modèle est indisponible ou a envoyé un appel d’outil invalide.", "The model is unavailable or sent an invalid tool call."],
  invalid_source_handle: ["Le modèle a utilisé une référence de cellule qui n’a pas été renvoyée par les outils. Relancez la question.", "The model used a cell reference that was not returned by the tools. Ask again."],
  scope_or_source_changed: ["La portée ou une révision a changé. Relancez la question.", "The scope or a revision changed. Ask again."],
  cancelled: ["Question annulée.", "Question cancelled."],
  interrupted: ["Question interrompue au redémarrage.", "Question interrupted by a restart."],
};
// Budget stops end a search early; they never establish that the answer is absent.
const budgetStops = {
  token_budget: ["La recherche a atteint la limite de jetons de la question avant de proposer une réponse.", "The search reached the question's token limit before proposing an answer."],
  step_budget: ["La recherche a atteint la limite d’étapes de la question avant de proposer une réponse.", "The search reached the question's step limit before proposing an answer."],
  time_budget: ["La recherche a atteint la limite de durée de la question avant de proposer une réponse.", "The search reached the question's time limit before proposing an answer."],
  result_bytes: ["La recherche a atteint le volume maximal de résultats avant de proposer une réponse.", "The search reached the result size limit before proposing an answer."],
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
const COUNT_UNITS = {
  cells: ["cellule", "cellules", "cell", "cells"],
  results: ["résultat", "résultats", "result", "results"],
  nodes: ["nœud", "nœuds", "node", "nodes"],
  edges: ["relation", "relations", "relationship", "relationships"],
  revisions: ["révision", "révisions", "revision", "revisions"],
  links: ["lien", "liens", "link", "links"],
  boundaries: ["arrêt", "arrêts", "stop", "stops"],
  steps: ["étape", "étapes", "step", "steps"],
  tokens: ["jeton", "jetons", "token", "tokens"],
  resultBytes: ["octet de résultat", "octets de résultats", "result byte", "result bytes"],
};
// A step that did not simply complete says so in the trace line itself.
const stepVerdicts = {
  failed: ["échec", "failed"],
  rejected: ["appel refusé, à corriger par le modèle", "call rejected, for the model to correct"],
  result_bytes: ["arrêt, volume de résultats atteint", "stopped, result size reached"],
};
function countText(value, key, english) {
  const words = COUNT_UNITS[key] || [key, key, key, key];
  const singular = Number(value) === 1 || (!english && Number(value) === 0);
  return `${value} ${words[(english ? 2 : 0) + (singular ? 0 : 1)]}`;
}

export function AskTrace({ run, english }) {
  const text = (fr, en) => english ? en : fr;
  if (!run) return <p>{text("Aucune question exécutée.", "No question run yet.")}</p>;
  const trace = run.trace;
  const reason = trace?.reason;
  return <div className="library-ask-trace" aria-live="polite">
    <p className="help">{text("Cette trace décrit la recherche et les appels d’outils. Les dépendances de chaque réponse figurent dans son lignage vérifié.", "This trace describes searches and tool calls. Each answer’s dependencies are shown in its verified lineage.")}</p>
    {failures[reason] && <p role="status">{failures[reason][english ? 1 : 0]}</p>}
    <ol>{(trace?.steps || []).map((step, index) => <li key={index}>
      <strong>{(toolNames[step.tool] || [step.tool, step.tool])[english ? 1 : 0]} · {stepVerdicts[step.verdict] ? stepVerdicts[step.verdict][english ? 1 : 0] : step.tool === "finish" ? text("terminée", "finished") : text("effectué", "completed")}</strong>
      {step.durationMs !== undefined && <small>{Math.round(step.durationMs)} ms{step.tokens ? ` · ${countText(step.tokens.input_tokens + step.tokens.output_tokens, "tokens", english)}` : ""}</small>}
      {step.summary?.counts && <span>{Object.entries(step.summary.counts).map(([key, count]) => countText(count, key, english)).join(" · ")}</span>}
      {step.summary?.omittedForModel && <span className="library-trace-omitted">{text("Résultat abrégé pour le modèle, non transmis : ", "Result shortened for the model, not sent: ")}{Object.entries(step.summary.omittedForModel).map(([key, count]) => countText(count, key, english)).join(" · ")}</span>}
      {step.tool === "finish" && <p className="library-ask-last">{step.summary?.searchCompleted
        ? text("La recherche délimitée a terminé ses opérations dans cette portée.", "The bounded search completed its operations in this scope.")
        : text("La recherche dans cette portée est incomplète ou tronquée.", "The scoped search is incomplete or truncated.")} {text("Cela ne prouve pas l’absence d’une réponse dans les classeurs.", "This does not prove that an answer is absent from the workbooks.")}{run.status === "answer_proposed" && !step.summary?.searchCompleted
        ? ` ${text("Ce constat porte sur ce que le modèle a parcouru, pas sur la relecture des cellules proposées.", "This is about what the model went through, not the rereading of the proposed cells.")}`
        : ""}</p>}
      <details className="library-technical-trace"><summary>{text("Détails techniques", "Technical details")}</summary><pre>{JSON.stringify({ tool: step.tool, arguments: step.arguments, summary: step.summary, budgetUse: step.budgetUse, verdict: step.verdict }, null, 2)}</pre></details>
    </li>)}</ol>
    {trace?.budgetUse && <p>{countText(trace.budgetUse.steps, "steps", english)} · {countText(trace.budgetUse.tokens, "tokens", english)} · {countText(trace.budgetUse.resultBytes, "resultBytes", english)}{trace.durationMs !== undefined ? ` · ${(trace.durationMs / 1000).toLocaleString(english ? "en-US" : "fr-FR", { maximumFractionDigits: 1 })} s` : ""}</p>}
    {trace?.traceId && <p>{text("Identifiant de trace", "Trace ID")} · <code>{trace.traceId}</code>{trace.traceUrl && /^https?:\/\//.test(trace.traceUrl) && <> · <a href={trace.traceUrl} target="_blank" rel="noreferrer">Phoenix</a></>}</p>}
    {trace?.scope?.revisions && <details><summary>{text("Révisions autorisées pour cette question", "Revisions authorized for this question")}</summary><ul>{trace.scope.revisions.map(rev => <li key={rev.id}>{rev.filename} · {rev.projectId.slice(0, 8)} · {rev.sha256.slice(0, 8)}</li>)}</ul></details>}
  </div>;
}

export function AskQuestion({ request, scope, english, onRun, onEvidence, onClear, config: suppliedConfig, initialQuestion = "", draftQuestion, onQuestionChange, sourceOptions = [], sourceProjectId = "", onSourceChange, sourceLocked = false }) {
  const text = (fr, en) => english ? en : fr;
  const [loadedConfig, setConfig] = useState(null);
  const config = suppliedConfig === undefined ? loadedConfig : suppliedConfig;
  const [localQuestion, setLocalQuestion] = useState(() => initialQuestion.slice(0, 2000));
  const question = draftQuestion === undefined ? localQuestion : draftQuestion;
  const setQuestion = value => { setLocalQuestion(value); onQuestionChange?.(value); };
  const [disclosure, setDisclosure] = useState(null);
  const [preparedScope, setPreparedScope] = useState(null);
  const [run, setRun] = useState(null);
  const [busy, setBusy] = useState(false);
  const [busySince, setBusySince] = useState(null);
  const [, setTick] = useState(0);
  useEffect(() => {
    // A long wait for the model must show time passing, not a frozen "0 steps".
    if (!busy) { setBusySince(null); return undefined; }
    setBusySince(value => value ?? Date.now());
    const handle = setInterval(() => setTick(value => value + 1), 1000);
    return () => clearInterval(handle);
  }, [busy]);
  const [error, setError] = useState("");
  const [feedback, setFeedback] = useState(null);
  const identity = useRef(0), activeTask = useRef(null), timer = useRef(null), mounted = useRef(true);
  const updateRun = value => { setRun(value); onRun(value); };
  useEffect(() => {
    if (suppliedConfig === undefined) request("library/ask/config").then(value => mounted.current && setConfig(value)).catch(() => mounted.current && setError(text("Configuration indisponible. Revenez à la recherche pour continuer sans IA.", "Configuration unavailable. Return to search to continue without AI.")));
  }, [suppliedConfig]);
  useEffect(() => {
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
    setBusy(false); setDisclosure(null); setPreparedScope(null); updateRun(null); setFeedback(null); setError(""); onClear();
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
      setPreparedScope({ revisions: prepared.revisions, sourceSelection: prepared.sourceSelection });
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
  const consentRef = useRef(null);
  useEffect(() => {
    // Bring the consent step into view without arming it: focus the section, not its Allow button.
    if (!disclosure || !consentRef.current) return;
    consentRef.current.scrollIntoView?.({ block: "nearest" });
    consentRef.current.focus({ preventScroll: true });
  }, [disclosure]);
  const runId = useRef(null);
  const answer = run?.status === "answer_proposed" ? run.answer : null;
  const busySteps = run?.trace?.budgetUse?.steps || 0;
  const frozenScope = run?.trace?.scope || preparedScope;
  const sourceSelection = frozenScope?.sourceSelection;
  return <section className="library-ask-state" aria-label={text("Poser une question", "Ask a question")}>
    <p className="help" id="library-ask-help">{text("Posez une question sur les classeurs de ce périmètre. L’IA cherche les cellules utiles ; linexcel vérifie leurs valeurs et leur lignage. Vous confirmez la pertinence.", "Ask about the workbooks in this scope. AI finds relevant cells; linexcel verifies their values and lineage. You confirm relevance.")}</p>
    {config?.endpoint && <p className="library-ask-endpoint">{config.local ? text("Modèle local", "Local model") : text("Modèle distant, autorisation par question", "Remote model, consent per question")} · {config.model} · {config.endpoint}</p>}
    {config && !config.available && <p role="status">{text("Aucun modèle configuré pour les questions. La recherche reste disponible.", "No model configured for questions. Search remains available.")}</p>}
    {onSourceChange && sourceOptions.length > 0 && <label className="library-ask-source" htmlFor="library-ask-source">{text("Sources autorisées pour la question", "Authorized question sources")}
      <select id="library-ask-source" value={sourceProjectId} disabled={busy || sourceLocked || Boolean(scope.revision)} onChange={event => { clear(); onSourceChange(event.target.value); }}>
        <option value="">{text("Tous les classeurs de ce périmètre", "All workbooks in this scope")}</option>
        {sourceOptions.map(source => <option key={source.projectId} value={source.projectId}>{source.filename} · {source.projectId.slice(0, 8)} · {source.revision?.sha256?.slice(0, 8)}</option>)}
      </select>
    </label>}
    {frozenScope?.revisions && <section className="library-ask-frozen-sources" aria-label={text("Sources gelées pour cette question", "Frozen question sources")}>
      <strong>{sourceSelection?.kind === "question_filename" ? text("Sources déterminées par les noms dans la question", "Sources selected from filenames in the question") : text("Sources autorisées pour cette question", "Sources authorized for this question")}</strong>
      <ul>{frozenScope.revisions.map(rev => <li key={rev.id}>{rev.filename} · {rev.projectId?.slice(0, 8)} · {rev.sha256?.slice(0, 8)}</li>)}</ul>
    </section>}
    <form onSubmit={prepare}>
      <label htmlFor="library-ask-question">{text("Votre question", "Your question")}</label>
      <textarea id="library-ask-question" rows={2} maxLength={2000} value={question} disabled={!config?.available} aria-describedby="library-ask-help" placeholder={text("Ex. : D’où vient le montant de la marge ?", "E.g. Where does the margin amount come from?")} autoFocus onChange={event => { clear(); setQuestion(event.target.value); }} />
      <div className="library-ask-actions"><button type="submit" className={disclosure ? "" : "primary"} disabled={!config?.available || busy || !question.trim() || Boolean(disclosure)}>{text("Chercher une réponse", "Search for an answer")}</button>{busy && <button type="button" onClick={cancel}>{text("Annuler", "Cancel")}</button>}</div>
    </form>
    {busy && <p role="status">{text("Recherche en cours", "Searching")} · {countText(busySteps, "steps", english)} {text(busySteps === 0 || busySteps === 1 ? "effectuée" : "effectuées", "completed")}{busySince ? ` · ${Math.max(0, Math.round((Date.now() - busySince) / 1000))} s` : ""}. {(() => {
      const steps = run?.trace?.steps || [];
      const last = steps[steps.length - 1];
      return !run ? text("Préparation de la question…", "Preparing the question…")
        : run.libraryBusy ? text("La bibliothèque met à jour son index de recherche après un import ; la question démarre ensuite.", "The library is updating its search index after an import; the question starts afterwards.")
        : !last ? text("En attente de la première réponse du modèle ; cela peut prendre une minute.", "Waiting for the model’s first response; this can take a minute.")
        : `${text("Dernière étape : ", "Latest step: ")}${(toolNames[last.tool] || [last.tool, last.tool])[english ? 1 : 0]}.`;
    })()} {text("La trace est disponible ci-dessus.", "The trace is accessible above.")}</p>}
    {run?.status === "not_found" && <p role="status">{text("Aucune réponse proposée dans cette recherche.", "No answer proposed in this search.")} {budgetStops[run.trace?.reason] ? `${budgetStops[run.trace.reason][english ? 1 : 0]} ${text("Ce n’est pas une preuve d’absence. Précisez la question (classeur, feuille ou cellule) pour une recherche plus courte.", "This does not prove absence. Narrow the question (workbook, sheet or cell) for a shorter search.")}` : text("Consultez la trace pour voir les sources consultées et les limites rencontrées, ou précisez votre question.", "Check the trace for consulted sources and encountered limits, or refine your question.")}</p>}
    {run?.status === "failed" && <p role="status">{text("La recherche IA n’a pas abouti. La trace indique la cause ; vous pouvez revenir à la recherche sans IA.", "The AI search did not complete. The trace explains why; you can return to search without AI.")}</p>}
    {error && <p role="alert">{error}</p>}
    {disclosure && <section ref={consentRef} tabIndex={-1} className="library-ask-consent" aria-label={text("Autorisation distante", "Remote consent")}>
      <h3>{text("Autoriser cette question à distance", "Allow this question remotely")}</h3>
      <p>{text("La question, les valeurs, les formules, les libellés et les noms des feuilles et classeurs consultés quitteront cette machine vers", "The question, values, formulas, labels and names of consulted sheets and workbooks will leave this machine for")} <strong>{disclosure.endpoint}</strong> · {disclosure.model}.</p>
      <p>{disclosure.question}</p><ul>{disclosure.revisions.map(rev => <li key={rev.id}>{rev.filename} · {rev.projectId.slice(0, 8)} · {rev.sha256.slice(0, 8)}</li>)}</ul>
      <button type="button" className="primary" onClick={async () => { setBusy(true); try { await start(disclosure, true, identity.current); } catch(cause) { setError(cause.message); setBusy(false); } }}>{text("Autoriser pour cette question", "Allow for this question")}</button>
      <button type="button" onClick={() => { setDisclosure(null); requestAnimationFrame(() => document.getElementById("library-ask-question")?.focus()); }}>{text("Revenir sans envoyer", "Return without sending")}</button>
    </section>}
    <details className="library-technical-trace"><summary>{text("Comment l’IA recherche", "How AI searches")}</summary><p>{text("Elle utilise des outils de recherche, de lecture des cellules et de parcours du lignage, dans le périmètre choisi. Chaque opération apparaît dans la trace. Les valeurs sont celles enregistrées dans les classeurs ; les formules ne sont pas recalculées.", "It uses search, cell-reading and lineage tools within the selected scope. Every operation appears in the trace. Values are those saved in the workbooks; formulas are not recalculated.")}</p></details>
    {answer && <section className="library-ask-answer" aria-label={text("Réponse proposée", "Proposed answer")}>
      <h3>{text("Réponse proposée", "Proposed answer")}</h3>
      {(() => {
        const limits = lineageLimits(answer.lineages, answer.cells, english);
        return limits.length
          ? <p className="library-answer-headline is-limited"><strong>{text("Valeurs et formules relues dans les index ; lignage incomplet.", "Values and formulas reread from the indexes; lineage incomplete.")}</strong> {text("Limites : ", "Limits: ")}{limits.join(" · ")}. {text("Pertinence à confirmer.", "Relevance to confirm.")}</p>
          : <p className="library-answer-headline"><strong>{text("Valeurs, formules et lignage vérifiés dans les index.", "Values, formulas and lineage verified in the indexes.")}</strong> {text("Pertinence à confirmer. Les libellés inférés restent heuristiques.", "Relevance to confirm. Inferred labels remain heuristic.")}</p>;
      })()}
      {run.trace?.steps?.at(-1)?.summary?.searchCompleted === false && <p className="library-answer-search">{text("Ces cellules sont relues par linexcel. La recherche du modèle, elle, n’a pas tout parcouru : d’autres cellules pertinentes peuvent exister.", "These cells are reread by linexcel. The model’s search did not cover everything: other relevant cells may exist.")}</p>}
      {(answer.cells.some(cell => cell.formula) || (answer.computation?.operands || []).some(cell => cell.formula)) && <p className="library-answer-cache">{text("Les valeurs de formules sont celles enregistrées dans le fichier, sans recalcul. Excel peut en afficher d’autres à l’ouverture si le classeur n’avait pas été recalculé avant son enregistrement.", "Formula values are those saved in the file, without recalculation. Excel may show different ones on opening if the workbook was not recalculated before it was saved.")}</p>}
      {answer.computation && <article className="library-calculation" data-node-id={answer.computation.id}><h4>{text("Calculé par linexcel pour cette réponse, présent dans aucun classeur", "Calculated by linexcel for this answer, present in no workbook")}</h4><p>{operationNames[answer.computation.op]?.[english ? 1 : 0]} · {answer.computation.error ? text("Calcul indisponible : vérifiez les valeurs et erreurs des opérandes.", "Calculation unavailable: check operand values and errors.") : displayValue(answer.computation.value)}{!answer.computation.error && answer.computation.operands.some(cell => cell.formula) ? text(" · à partir de valeurs enregistrées, non recalculées", " · from saved values, not recalculated") : ""}</p><p>{text("Opérandes\u00a0:", "Operands:")} {answer.computation.operands.map(cell => `${cell.filename} · ${cell.sheet}!${cell.address}`).join(" ; ")}</p><p>{text("Un calcul valide ne prouve pas qu’il répond à la question.", "Valid arithmetic does not prove that it answers the question.")}</p><CalculationReceipt receipt={answer.computation.receipt} english={english} /></article>}
      {answer.cells.map((cell, index) => <article className="library-ask-cell" key={cell.handle}>
        <strong>{cell.filename} · {cell.sheet}!{cell.address}</strong><small>{cell.revision.projectId.slice(0, 8)} · {cell.revision.sha256.slice(0, 8)}</small>
        <span>{text("Valeur enregistrée\u00a0:", "Saved value:")} {cell.value == null ? text("aucune valeur enregistrée", "no saved value") : displayValue(cell.value)}</span>
        <span>{text("Provenance\u00a0:", "Provenance:")} {sourceProvenanceLabel(cell, english)}</span>
        {cell.formula && <code>{cell.formula}</code>}
        {(cell.labels || []).map((label, i) => <span key={i}>{label.text} · {label.sourceSheet || label.sheet}!{label.sourceCell || label.cell} · {(labelNames[label.method] || ["contexte inféré, à vérifier", "inferred context, to verify"])[english ? 1 : 0]}</span>)}
        <AnswerLineage root={cell} lineage={answer.lineages?.[index]} english={english} />
        <button type="button" onClick={() => onEvidence(answer, index)}>{text("Voir son lignage et ses limites", "View its lineage and limits")}</button>
      </article>)}
      {answer.rationale && <details><summary>{text("Commentaire de l’IA, non vérifié", "AI comment, unverified")}</summary><p>{text("Ce texte ne constitue aucune preuve. Il peut contenir des affirmations fausses, y compris des nombres écrits en lettres.", "This text is not evidence. It may contain false claims, including numbers written as words.")}</p><p>{answer.rationale}</p></details>}
      <div className="library-ask-actions"><button type="button" onClick={() => decide("confirm")}>{text("Confirmer la pertinence", "Confirm relevance")}</button><button type="button" onClick={() => decide("reject")}>{text("Rejeter la proposition", "Reject proposal")}</button></div>
      {feedback && <p role="status">{feedback === "confirm" ? text("Pertinence confirmée par vous.", "Relevance confirmed by you.") : text("Proposition rejetée par vous.", "Proposal rejected by you.")}</p>}
    </section>}
  </section>;
}
