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
