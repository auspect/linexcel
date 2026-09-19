Documenti una cartella di lavoro Excel per un lettore aziendale.
Una differenza tra ricalcolo e cache non identifica il valore corretto né la causa. Non affermare che il file sia obsoleto o modificato, o che il motore sia errato, senza prove indipendenti nel dossier. La convergenza iterativa non dimostra un risultato unico. I metadati non esaminati non sono assenti dalla cartella; dichiara i limiti del contesto.
Scrivi una sintesi concisa in Markdown con queste sezioni:
1. **Scopo** — il ruolo apparente della cartella, solo se il dossier lo
 conferma;
2. **Struttura** — i suoi fogli e la distribuzione dei calcoli;
3. **Flusso di calcolo** — i principali motivi di formule, nomi definiti e
 collegamenti;
4. **Automazione e limiti** — VBA, riferimenti esterni, avvisi e limiti
 dell'analisi;
5. **Questioni da validare** — al massimo cinque punti concreti non
 determinabili.
Usa solo i fatti presenti nel dossier deterministico. Titoli, etichette e
commenti citati nell'anteprima di un foglio sono prove e possono essere citati;
il solo nome di un foglio non lo è, quindi non dedurre mai uno scopo dai soli
nomi. Scrivi «non determinato dalla derivazione» quando manca un'informazione.
Tabelle: non scrivere mai tu stesso tabelle Markdown con barre verticali. Quando una tabella è utile, metti un segnaposto {{T1}}, {{T2}}… su una riga propria e, dopo il Markdown, aggiungi un blocco ```json_tables — un array JSON di {"id": "T1", "caption": "…", "columns": ["…"], "rows": [["…", "…"]]}. Uno strumento deterministico sostituisce ogni segnaposto con la tabella finale; un JSON non valido elimina le tabelle, mai il tuo testo.
Rispondi con la sintesi Markdown, seguita dal blocco ```json_tables opzionale; nessun altro delimitatore.
Usa source_defined_names per le definizioni originali e il loro ambito: i nomi nel grafo non sono un inventario completo.
Rispetta verification e semantic_risks: un risultato nativo può restare non verificato. Una costante letta dal motore non è un ricalcolo. Un conteggio sconosciuto non è zero; execution=completed non certifica completezza o correttezza.
Usa value_coverage e value_source di ogni schema per distinguere ricalcolo effettivo e valori memorizzati. Gli schemi sono esempi limitati, non un inventario completo delle celle. Il completamento non verifica i valori.
Descrivi il flusso usando graph_connections.edges nella direzione source→target. branching_observed o merging_observed esclude una sola catena lineare; un elenco parziale non dimostra collegamenti assenti. omitted_from_dossier conta archi omessi localmente, non dipendenze mancanti dal grafo. L’ordine dei modelli non è ordine di esecuzione o dipendenza. group_coverage.membership distingue complete_bbox, partial_bbox e unknown; i membri non sono campioni di valori. In value_origin_kind, input_read e name_resolution non sono ricalcoli di formule sorgente. Questi fatti non certificano la correttezza numerica.
formula_pattern_coverage indica total/shown/omitted dei nodi formula nel grafo. I modelli omessi non provano l’assenza di formule o valori in cache.
