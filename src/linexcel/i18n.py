"""Supported languages for the viewer UI and the AI documentation.

The set is deliberately **closed**. ``language`` selects a stored system prompt
and is interpolated into the generated viewer, so it is validated against this
allowlist rather than escaped or passed through: an arbitrary string would let a
caller steer the model's instructions or reach the generated JavaScript.

Adding a language means adding an entry here *and* in :mod:`linexcel.aidoc`'s
three prompt registries (node, workbook, screenshot). ``tests/test_lineage.py``
asserts they all stay in sync, so a partial addition fails the suite instead of
surfacing as raw keys in the report or a ``KeyError`` at generation time.

Provenance: ``en`` and ``fr`` were written by hand. The other seven languages,
here and in the prompt registries, were produced with AI assistance and have not
been reviewed by native speakers — corrections welcome.
"""

from __future__ import annotations

from typing import Any

#: Languages accepted by ``language=`` across the public API.
LANGUAGES = ("en", "fr", "es", "de", "it", "pt", "nl", "ja", "zh")

DEFAULT_LANGUAGE = "en"

#: Viewer interface strings. Every language carries the same key set; ``{name}``
#: placeholders are substituted client-side and must be preserved verbatim.
UI_STRINGS: dict[str, dict[str, str]] = {
    "en": {
        "graph": "Graph",
        "overview": "Workbook overview",
        "visual": "Visual preview",
        "search": "Search…",
        "search_label": "Search cells and formulas (Enter)",
        # Label-first so a single hit does not read "1 matches": the count is
        # now on screen, where the plural agreement is visible.
        "search_matches": "Matches: {count}",
        "search_none": "No matches",
        "search_clear": "Clear search",
        "theme_dark": "Dark theme",
        "fit_all": "Fit",
        "fit_sel": "Fit Selection",
        "zoom_in": "Zoom In",
        "zoom_out": "Zoom Out",
        "layout_flow": "Flow",
        "layout_organic": "Organic",
        "close": "Close",
        "formula": "Formula",
        "stretched_pattern": "Stretched pattern over {count} cells ({bbox}).",
        "value_heading": "Value",
        "value_samples": "Value samples",
        "value_from_file": "Read from Excel file",
        "value_recalc": "Recalculated by linexcel",
        "value_fallback": "linexcel fallback (error-guarded)",
        "value_recalc_desc": (
            "linexcel recomputed this value from the workbook's own inputs."
        ),
        "value_from_file_desc": (
            "Shown as Excel stored it in the file; linexcel did not recalculate it."
        ),
        "value_fallback_desc": (
            "linexcel could not complete this evaluation: the figure shown is an "
            "error-guarded fallback."
        ),
        "value_volatile": "Not recalculated (volatile)",
        "value_volatile_desc": (
            "TODAY, NOW and RAND answer differently every time they are computed, "
            "so linexcel keeps what the file stores rather than inventing a "
            "value the workbook never had."
        ),
        "external_books": "External workbooks",
        "external_read_folder": "read from the folder provided",
        "external_read_cache": "not read — value cached in this file",
        "external_read_none": "not read",
        "value_external": "Read from another workbook",
        "value_external_desc": (
            "This value comes from a workbook linexcel was pointed at, not from "
            "the analyzed file: the reference was replaced by the value before "
            "the formula was evaluated."
        ),
        "external_hint": (
            "Pass the folder holding these files (--refs-dir) so linexcel can "
            "read them."
        ),
        "query_source": "M source",
        "query_loaded": "Loaded into",
        "query_not_loaded": "Not loaded onto a sheet (connection only)",
        "query_reads": "Reads",
        "query_reads_hint": (
            "A source outside this workbook is named by the query, not read by "
            "linexcel: what it holds is not in the file."
        ),
        "value_col_file": "Excel file",
        "value_col_calc": "linexcel recalc",
        "value_col_cell": "Cell",
        "value_not_in_file": "Not stored",
        "value_not_recalc": "Not recalculated",
        "value_no_cache_desc": (
            "The file stores no value to compare with: Excel writes one only "
            "when it saves a calculated workbook."
        ),
        "sampled_cells": "{shown} cells sampled out of {count}.",
        "value_match": "The recalculated value matches the file",
        "value_mismatch": "The recalculated value differs from the file",
        "target": "Target",
        "step_decomp": "Step-by-step decomposition",
        "step_hint": "Each function/operator is evaluated individually.",
        "final_result": "Final result",
        "not_evaluated": "not evaluated",
        "precedents": "Precedents",
        "dependents": "Dependents",
        "cells": "cells",
        "ai_doc": "🤖 AI Documentation (Generated)",
        "ai_vision": "🤖 AI — read from the screenshot",
        "ai_overview": "🤖 AI Generated Overview",
        "ai_overview_desc": (
            "This overview was written by an AI model from the deterministic "
            "lineage. The facts presented are derived from the workbook's "
            "formulas and data."
        ),
        "fallback": (
            "Cytoscape could not be loaded (CDN access required). The JSON "
            "graph remains available via result.to_dict()."
        ),
        "stats": "{formulas} formulas · {nodes} nodes · {edges} edges{vba}",
        "kind_cell": "Formula",
        "kind_group": "Stretched formulas",
        "kind_input": "Source data",
        "kind_name": "Named cell/range",
        "kind_vba": "VBA",
        "kind_query": "Power Query",
        "kind_misc": "Other (aggregated)",
        "kind_opaque": "External reference",
        "placeholder_title": "Select a node",
        "placeholder_desc": (
            "Select a node in the graph to inspect its formula, computed "
            "value, step-by-step evaluation, and AI-generated documentation."
        ),
        "sheets_tab": "Sheets",
        "all_sheets": "All sheets",
        "sheet_filter": "Sheet filter",
        "sheet_dims": "{rows} rows × {cols} columns",
        "visibility": "Visibility",
        "freeze_panes": "Freeze panes",
        "hidden_columns": "Hidden columns",
        "merged_ranges": "Merged ranges",
        "comments": "Comments",
        "sheet_preview": "Cell preview",
        "sheet_render": "Rendered sheet",
        "details_panel": "Node details",
        "page": "Page {n}",
    },
    "fr": {
        "graph": "Graphe",
        "overview": "Synthèse générale",
        "visual": "Aperçu visuel",
        "search": "Rechercher…",
        "search_label": "Rechercher une cellule ou une formule (Entrée)",
        "search_matches": "Résultats : {count}",
        "search_none": "Aucun résultat",
        "search_clear": "Effacer la recherche",
        "theme_dark": "Thème sombre",
        "fit_all": "Ajuster",
        "fit_sel": "Ajuster la sélection",
        "zoom_in": "Zoom avant",
        "zoom_out": "Zoom arrière",
        "layout_flow": "Flux",
        "layout_organic": "Organique",
        "close": "Fermer",
        "formula": "Formule",
        "stretched_pattern": "Formule étirée sur {count} cellules ({bbox}).",
        "value_heading": "Valeur",
        "value_samples": "Échantillons de valeurs",
        "value_from_file": "Lue du fichier Excel",
        "value_recalc": "Recalculée par linexcel",
        "value_fallback": "repli linexcel (erreur protégée)",
        "value_recalc_desc": (
            "linexcel a recalculé cette valeur à partir des données du classeur."
        ),
        "value_from_file_desc": (
            "Valeur telle qu'Excel l'a enregistrée dans le fichier ; linexcel "
            "ne l'a pas recalculée."
        ),
        "value_fallback_desc": (
            "linexcel n'a pas pu mener l'évaluation à son terme : la valeur "
            "affichée est un repli protégé contre l'erreur."
        ),
        "value_volatile": "Non recalculée (volatile)",
        "value_volatile_desc": (
            "TODAY, NOW et RAND répondent différemment à chaque calcul : linexcel "
            "conserve donc la valeur du fichier plutôt que d'en inventer une "
            "que le classeur n'a jamais eue."
        ),
        "external_books": "Classeurs externes",
        "external_read_folder": "lu depuis le dossier fourni",
        "external_read_cache": ("non lu — valeur mise en cache dans ce fichier"),
        "external_read_none": "non lu",
        "value_external": "Lue d'un autre classeur",
        "value_external_desc": (
            "Cette valeur provient d'un classeur vers lequel linexcel a été "
            "pointé, pas du fichier analysé : la référence a été remplacée par "
            "la valeur avant l'évaluation de la formule."
        ),
        "external_hint": (
            "Indiquez le dossier contenant ces fichiers (--refs-dir) pour que "
            "linexcel puisse les lire."
        ),
        "query_source": "Source M",
        "query_loaded": "Chargée dans",
        "query_not_loaded": "Non chargée dans une feuille (connexion seule)",
        "query_reads": "Lit",
        "query_reads_hint": (
            "Une source hors de ce classeur est nommée par la requête, pas lue "
            "par linexcel : son contenu n'est pas dans le fichier."
        ),
        "value_col_file": "Fichier Excel",
        "value_col_calc": "Recalcul linexcel",
        "value_col_cell": "Cellule",
        "value_not_in_file": "Absente",
        "value_not_recalc": "Non recalculée",
        "value_no_cache_desc": (
            "Le fichier ne contient aucune valeur à comparer : Excel n'en "
            "enregistre une qu'en sauvegardant un classeur calculé."
        ),
        "sampled_cells": "{shown} cellules échantillonnées sur {count}.",
        "value_match": "La valeur recalculée correspond au fichier",
        "value_mismatch": "La valeur recalculée diffère du fichier",
        "target": "Cible",
        "step_decomp": "Décomposition pas-à-pas",
        "step_hint": "Chaque fonction et opérateur est évalué individuellement.",
        "final_result": "Résultat final",
        "not_evaluated": "non évalué",
        "precedents": "Précédents",
        "dependents": "Dépendants",
        "cells": "cellules",
        "ai_doc": "🤖 Documentation IA (Générée)",
        "ai_vision": "🤖 IA — lu sur la capture",
        "ai_overview": "🤖 Synthèse Générée par IA",
        "ai_overview_desc": (
            "Cette synthèse a été rédigée par un modèle d'IA à partir du "
            "lignage de calculs déterministe. Les faits présentés proviennent "
            "des formules et des données du classeur."
        ),
        "fallback": (
            "Cytoscape n'a pas pu être chargé (accès CDN requis). Le graphe "
            "JSON reste disponible via result.to_dict()."
        ),
        "stats": "{formulas} formules · {nodes} nœuds · {edges} liens{vba}",
        "kind_cell": "Formule",
        "kind_group": "Formules étirées",
        "kind_input": "Source de données",
        "kind_name": "Cellule/plage nommée",
        "kind_vba": "VBA",
        "kind_query": "Power Query",
        "kind_misc": "Autre (agrégé)",
        "kind_opaque": "Référence externe",
        "placeholder_title": "Sélectionner un nœud",
        "placeholder_desc": (
            "Sélectionnez un nœud dans le graphe pour afficher sa formule, sa "
            "valeur calculée, sa décomposition pas à pas et sa documentation IA."
        ),
        "sheets_tab": "Feuilles",
        "all_sheets": "Toutes les feuilles",
        "sheet_filter": "Filtre par feuille",
        "sheet_dims": "{rows} lignes × {cols} colonnes",
        "visibility": "Visibilité",
        "freeze_panes": "Volets figés",
        "hidden_columns": "Colonnes masquées",
        "merged_ranges": "Plages fusionnées",
        "comments": "Commentaires",
        "sheet_preview": "Aperçu des cellules",
        "sheet_render": "Rendu de la feuille",
        "details_panel": "Détails du nœud",
        "page": "Page {n}",
    },
    "es": {
        "graph": "Grafo",
        "overview": "Resumen del libro",
        "visual": "Vista previa visual",
        "search": "Buscar…",
        "search_label": "Buscar celdas y fórmulas (Intro)",
        "search_matches": "Resultados: {count}",
        "search_none": "Sin resultados",
        "search_clear": "Borrar la búsqueda",
        "theme_dark": "Tema oscuro",
        "fit_all": "Ajustar",
        "fit_sel": "Ajustar selección",
        "zoom_in": "Acercar",
        "zoom_out": "Alejar",
        "layout_flow": "Flujo",
        "layout_organic": "Orgánico",
        "close": "Cerrar",
        "formula": "Fórmula",
        "stretched_pattern": "Patrón extendido sobre {count} celdas ({bbox}).",
        "value_heading": "Valor",
        "value_samples": "Muestras de valores",
        "value_from_file": "Leído del archivo Excel",
        "value_recalc": "Recalculado por linexcel",
        "value_fallback": "valor alternativo de linexcel (error protegido)",
        "value_recalc_desc": (
            "linexcel ha recalculado este valor a partir de los datos del libro."
        ),
        "value_from_file_desc": (
            "Valor tal como Excel lo guardó en el archivo; linexcel no lo ha "
            "recalculado."
        ),
        "value_fallback_desc": (
            "linexcel no ha podido completar esta evaluación: la cifra mostrada "
            "es un valor alternativo protegido frente a errores."
        ),
        "value_volatile": "No recalculado (volátil)",
        "value_volatile_desc": (
            "TODAY, NOW y RAND responden de forma distinta en cada cálculo, así que "
            "linexcel conserva lo que guarda el archivo en lugar de inventar un "
            "valor que el libro nunca tuvo."
        ),
        "external_books": "Libros externos",
        "external_read_folder": "leído de la carpeta indicada",
        "external_read_cache": "no leído — valor en caché en este archivo",
        "external_read_none": "no leído",
        "value_external": "Leído de otro libro",
        "value_external_desc": (
            "Este valor procede de un libro al que se apuntó a linexcel, no del "
            "archivo analizado: la referencia se sustituyó por el valor antes "
            "de evaluar la fórmula."
        ),
        "external_hint": (
            "Indique la carpeta que contiene estos archivos (--refs-dir) para "
            "que linexcel pueda leerlos."
        ),
        "query_source": "Origen M",
        "query_loaded": "Cargada en",
        "query_not_loaded": "No cargada en una hoja (solo conexión)",
        "query_reads": "Lee",
        "query_reads_hint": (
            "La consulta nombra un origen externo a este libro, pero linexcel "
            "no lo lee: su contenido no está en el archivo."
        ),
        "value_col_file": "Archivo Excel",
        "value_col_calc": "Recálculo linexcel",
        "value_col_cell": "Celda",
        "value_not_in_file": "No guardado",
        "value_not_recalc": "No recalculado",
        "value_no_cache_desc": (
            "El archivo no guarda ningún valor con el que comparar: Excel solo "
            "lo escribe al guardar un libro calculado."
        ),
        "sampled_cells": "{shown} celdas muestreadas de {count}.",
        "value_match": "El valor recalculado coincide con el archivo",
        "value_mismatch": "El valor recalculado difiere del archivo",
        "target": "Destino",
        "step_decomp": "Descomposición paso a paso",
        "step_hint": "Cada función y operador se evalúa individualmente.",
        "final_result": "Resultado final",
        "not_evaluated": "no evaluado",
        "precedents": "Precedentes",
        "dependents": "Dependientes",
        "cells": "celdas",
        "ai_doc": "🤖 Documentación IA (generada)",
        "ai_vision": "🤖 IA — leído en la captura",
        "ai_overview": "🤖 Resumen generado por IA",
        "ai_overview_desc": (
            "Este resumen ha sido redactado por un modelo de IA a partir del "
            "linaje de cálculos determinista. Los hechos presentados proceden "
            "de las fórmulas y los datos del libro."
        ),
        "fallback": (
            "No se pudo cargar Cytoscape (se requiere acceso al CDN). El grafo "
            "JSON sigue disponible mediante result.to_dict()."
        ),
        "stats": "{formulas} fórmulas · {nodes} nodos · {edges} enlaces{vba}",
        "kind_cell": "Fórmula",
        "kind_group": "Fórmulas extendidas",
        "kind_input": "Datos de origen",
        "kind_name": "Celda/rango con nombre",
        "kind_vba": "VBA",
        "kind_query": "Power Query",
        "kind_misc": "Otros (agregados)",
        "kind_opaque": "Referencia externa",
        "placeholder_title": "Seleccione un nodo",
        "placeholder_desc": (
            "Seleccione un nodo del grafo para consultar su fórmula, su valor "
            "calculado, su evaluación paso a paso y su documentación generada "
            "por IA."
        ),
        "sheets_tab": "Hojas",
        "all_sheets": "Todas las hojas",
        "sheet_filter": "Filtro por hoja",
        "sheet_dims": "{rows} filas × {cols} columnas",
        "visibility": "Visibilidad",
        "freeze_panes": "Paneles inmovilizados",
        "hidden_columns": "Columnas ocultas",
        "merged_ranges": "Rangos combinados",
        "comments": "Comentarios",
        "sheet_preview": "Vista previa de celdas",
        "sheet_render": "Hoja renderizada",
        "details_panel": "Detalles del nodo",
        "page": "Página {n}",
    },
    "de": {
        "graph": "Graph",
        "overview": "Arbeitsmappen-Überblick",
        "visual": "Visuelle Vorschau",
        "search": "Suchen…",
        "search_label": "Zellen und Formeln durchsuchen (Eingabetaste)",
        "search_matches": "Treffer: {count}",
        "search_none": "Keine Treffer",
        "search_clear": "Suche löschen",
        "theme_dark": "Dunkles Design",
        "fit_all": "Einpassen",
        "fit_sel": "Auswahl einpassen",
        "zoom_in": "Vergrößern",
        "zoom_out": "Verkleinern",
        "layout_flow": "Fluss",
        "layout_organic": "Organisch",
        "close": "Schließen",
        "formula": "Formel",
        "stretched_pattern": "Gezogenes Muster über {count} Zellen ({bbox}).",
        "value_heading": "Wert",
        "value_samples": "Wertebeispiele",
        "value_from_file": "Aus der Excel-Datei gelesen",
        "value_recalc": "Von linexcel neu berechnet",
        "value_fallback": "linexcel-Ersatzwert (fehlergeschützt)",
        "value_recalc_desc": (
            "linexcel hat diesen Wert aus den Daten der Arbeitsmappe neu berechnet."
        ),
        "value_from_file_desc": (
            "Wert so, wie Excel ihn in der Datei gespeichert hat; linexcel hat "
            "ihn nicht neu berechnet."
        ),
        "value_fallback_desc": (
            "linexcel konnte diese Auswertung nicht abschließen: der angezeigte "
            "Wert ist ein fehlergeschützter Ersatzwert."
        ),
        "value_volatile": "Nicht neu berechnet (volatil)",
        "value_volatile_desc": (
            "TODAY, NOW und RAND antworten bei jeder Berechnung anders, daher behält "
            "linexcel den Wert aus der Datei, statt einen zu erfinden, den die "
            "Arbeitsmappe nie hatte."
        ),
        "external_books": "Externe Arbeitsmappen",
        "external_read_folder": "aus dem angegebenen Ordner gelesen",
        "external_read_cache": ("nicht gelesen — Wert in dieser Datei gespeichert"),
        "external_read_none": "nicht gelesen",
        "value_external": "Aus einer anderen Arbeitsmappe gelesen",
        "value_external_desc": (
            "Dieser Wert stammt aus einer Arbeitsmappe, auf die linexcel "
            "verwiesen wurde, nicht aus der analysierten Datei: der Bezug wurde "
            "vor der Auswertung durch den Wert ersetzt."
        ),
        "external_hint": (
            "Geben Sie den Ordner mit diesen Dateien an (--refs-dir), damit "
            "linexcel sie lesen kann."
        ),
        "query_source": "M-Quelltext",
        "query_loaded": "Geladen in",
        "query_not_loaded": "Nicht in ein Blatt geladen (nur Verbindung)",
        "query_reads": "Liest",
        "query_reads_hint": (
            "Eine Quelle außerhalb dieser Arbeitsmappe wird von der Abfrage "
            "genannt, aber nicht von linexcel gelesen: ihr Inhalt steht nicht "
            "in der Datei."
        ),
        "value_col_file": "Excel-Datei",
        "value_col_calc": "linexcel-Neuberechnung",
        "value_col_cell": "Zelle",
        "value_not_in_file": "Nicht gespeichert",
        "value_not_recalc": "Nicht neu berechnet",
        "value_no_cache_desc": (
            "Die Datei enthält keinen Wert zum Vergleich: Excel schreibt ihn nur "
            "beim Speichern einer berechneten Arbeitsmappe."
        ),
        "sampled_cells": "{shown} von {count} Zellen als Stichprobe.",
        "value_match": "Der neu berechnete Wert stimmt mit der Datei überein",
        "value_mismatch": "Der neu berechnete Wert weicht von der Datei ab",
        "target": "Ziel",
        "step_decomp": "Schrittweise Zerlegung",
        "step_hint": "Jede Funktion und jeder Operator wird einzeln ausgewertet.",
        "final_result": "Endergebnis",
        "not_evaluated": "nicht ausgewertet",
        "precedents": "Vorgänger",
        "dependents": "Nachfolger",
        "cells": "Zellen",
        "ai_doc": "🤖 KI-Dokumentation (generiert)",
        "ai_vision": "🤖 KI — aus dem Screenshot gelesen",
        "ai_overview": "🤖 KI-generierter Überblick",
        "ai_overview_desc": (
            "Dieser Überblick wurde von einem KI-Modell auf Basis der "
            "deterministischen Berechnungsherkunft verfasst. Die dargestellten "
            "Fakten stammen aus den Formeln und Daten der Arbeitsmappe."
        ),
        "fallback": (
            "Cytoscape konnte nicht geladen werden (CDN-Zugriff erforderlich). "
            "Der JSON-Graph bleibt über result.to_dict() verfügbar."
        ),
        "stats": "{formulas} Formeln · {nodes} Knoten · {edges} Kanten{vba}",
        "kind_cell": "Formel",
        "kind_group": "Gezogene Formeln",
        "kind_input": "Quelldaten",
        "kind_name": "Benannte Zelle/Bereich",
        "kind_vba": "VBA",
        "kind_query": "Power Query",
        "kind_misc": "Sonstige (aggregiert)",
        "kind_opaque": "Externer Bezug",
        "placeholder_title": "Knoten auswählen",
        "placeholder_desc": (
            "Wählen Sie einen Knoten im Graphen aus, um seine Formel, seinen "
            "berechneten Wert, seine schrittweise Auswertung und seine "
            "KI-generierte Dokumentation anzuzeigen."
        ),
        "sheets_tab": "Blätter",
        "all_sheets": "Alle Blätter",
        "sheet_filter": "Blattfilter",
        "sheet_dims": "{rows} Zeilen × {cols} Spalten",
        "visibility": "Sichtbarkeit",
        "freeze_panes": "Fixierte Fenster",
        "hidden_columns": "Ausgeblendete Spalten",
        "merged_ranges": "Verbundene Bereiche",
        "comments": "Kommentare",
        "sheet_preview": "Zellvorschau",
        "sheet_render": "Gerendertes Blatt",
        "details_panel": "Knotendetails",
        "page": "Seite {n}",
    },
    "it": {
        "graph": "Grafo",
        "overview": "Panoramica della cartella",
        "visual": "Anteprima visiva",
        "search": "Cerca…",
        "search_label": "Cerca celle e formule (Invio)",
        "search_matches": "Risultati: {count}",
        "search_none": "Nessun risultato",
        "search_clear": "Cancella la ricerca",
        "theme_dark": "Tema scuro",
        "fit_all": "Adatta",
        "fit_sel": "Adatta selezione",
        "zoom_in": "Ingrandisci",
        "zoom_out": "Riduci",
        "layout_flow": "Flusso",
        "layout_organic": "Organico",
        "close": "Chiudi",
        "formula": "Formula",
        "stretched_pattern": "Motivo esteso su {count} celle ({bbox}).",
        "value_heading": "Valore",
        "value_samples": "Campioni di valori",
        "value_from_file": "Letto dal file Excel",
        "value_recalc": "Ricalcolato da linexcel",
        "value_fallback": "valore di ripiego di linexcel (errore protetto)",
        "value_recalc_desc": (
            "linexcel ha ricalcolato questo valore a partire dai dati della "
            "cartella di lavoro."
        ),
        "value_from_file_desc": (
            "Valore così come Excel lo ha salvato nel file; linexcel non lo ha "
            "ricalcolato."
        ),
        "value_fallback_desc": (
            "linexcel non è riuscito a completare questa valutazione: il valore "
            "mostrato è un ripiego protetto dagli errori."
        ),
        "value_volatile": "Non ricalcolato (volatile)",
        "value_volatile_desc": (
            "TODAY, NOW e RAND rispondono in modo diverso a ogni calcolo, quindi "
            "linexcel conserva quanto è salvato nel file invece di inventare un "
            "valore che la cartella di lavoro non ha mai avuto."
        ),
        "external_books": "Cartelle di lavoro esterne",
        "external_read_folder": "letta dalla cartella indicata",
        "external_read_cache": ("non letta — valore memorizzato in questo file"),
        "external_read_none": "non letta",
        "value_external": "Letto da un'altra cartella di lavoro",
        "value_external_desc": (
            "Questo valore proviene da una cartella di lavoro indicata a "
            "linexcel, non dal file analizzato: il riferimento è stato "
            "sostituito dal valore prima di valutare la formula."
        ),
        "external_hint": (
            "Indica la cartella che contiene questi file (--refs-dir) perché "
            "linexcel possa leggerli."
        ),
        "query_source": "Sorgente M",
        "query_loaded": "Caricata in",
        "query_not_loaded": "Non caricata in un foglio (solo connessione)",
        "query_reads": "Legge",
        "query_reads_hint": (
            "Una sorgente esterna a questa cartella di lavoro è nominata dalla "
            "query, non letta da linexcel: il suo contenuto non è nel file."
        ),
        "value_col_file": "File Excel",
        "value_col_calc": "Ricalcolo linexcel",
        "value_col_cell": "Cella",
        "value_not_in_file": "Non salvato",
        "value_not_recalc": "Non ricalcolato",
        "value_no_cache_desc": (
            "Il file non contiene alcun valore da confrontare: Excel lo scrive "
            "solo salvando una cartella di lavoro calcolata."
        ),
        "sampled_cells": "{shown} celle campionate su {count}.",
        "value_match": "Il valore ricalcolato coincide con il file",
        "value_mismatch": "Il valore ricalcolato differisce dal file",
        "target": "Destinazione",
        "step_decomp": "Scomposizione passo passo",
        "step_hint": "Ogni funzione e operatore viene valutato singolarmente.",
        "final_result": "Risultato finale",
        "not_evaluated": "non valutato",
        "precedents": "Precedenti",
        "dependents": "Dipendenti",
        "cells": "celle",
        "ai_doc": "🤖 Documentazione IA (generata)",
        "ai_vision": "🤖 IA — letto dalla schermata",
        "ai_overview": "🤖 Panoramica generata dall'IA",
        "ai_overview_desc": (
            "Questa panoramica è stata redatta da un modello di IA a partire "
            "dalla derivazione deterministica dei calcoli. I fatti presentati "
            "provengono dalle formule e dai dati della cartella di lavoro."
        ),
        "fallback": (
            "Impossibile caricare Cytoscape (è necessario l'accesso al CDN). "
            "Il grafo JSON resta disponibile tramite result.to_dict()."
        ),
        "stats": "{formulas} formule · {nodes} nodi · {edges} archi{vba}",
        "kind_cell": "Formula",
        "kind_group": "Formule estese",
        "kind_input": "Dati di origine",
        "kind_name": "Cella/intervallo denominato",
        "kind_vba": "VBA",
        "kind_query": "Power Query",
        "kind_misc": "Altro (aggregato)",
        "kind_opaque": "Riferimento esterno",
        "placeholder_title": "Seleziona un nodo",
        "placeholder_desc": (
            "Seleziona un nodo nel grafo per consultarne la formula, il valore "
            "calcolato, la valutazione passo passo e la documentazione "
            "generata dall'IA."
        ),
        "sheets_tab": "Fogli",
        "all_sheets": "Tutti i fogli",
        "sheet_filter": "Filtro per foglio",
        "sheet_dims": "{rows} righe × {cols} colonne",
        "visibility": "Visibilità",
        "freeze_panes": "Riquadri bloccati",
        "hidden_columns": "Colonne nascoste",
        "merged_ranges": "Intervalli uniti",
        "comments": "Commenti",
        "sheet_preview": "Anteprima celle",
        "sheet_render": "Foglio renderizzato",
        "details_panel": "Dettagli del nodo",
        "page": "Pagina {n}",
    },
    "pt": {
        "graph": "Grafo",
        "overview": "Visão geral da pasta",
        "visual": "Pré-visualização",
        "search": "Pesquisar…",
        "search_label": "Pesquisar células e fórmulas (Enter)",
        "search_matches": "Resultados: {count}",
        "search_none": "Sem resultados",
        "search_clear": "Limpar a pesquisa",
        "theme_dark": "Tema escuro",
        "fit_all": "Ajustar",
        "fit_sel": "Ajustar seleção",
        "zoom_in": "Ampliar",
        "zoom_out": "Reduzir",
        "layout_flow": "Fluxo",
        "layout_organic": "Orgânico",
        "close": "Fechar",
        "formula": "Fórmula",
        "stretched_pattern": "Padrão estendido por {count} células ({bbox}).",
        "value_heading": "Valor",
        "value_samples": "Amostras de valores",
        "value_from_file": "Lido do ficheiro Excel",
        "value_recalc": "Recalculado pelo linexcel",
        "value_fallback": "valor alternativo do linexcel (erro protegido)",
        "value_recalc_desc": (
            "O linexcel recalculou este valor a partir dos dados da pasta de trabalho."
        ),
        "value_from_file_desc": (
            "Valor tal como o Excel o guardou no ficheiro; o linexcel não o recalculou."
        ),
        "value_fallback_desc": (
            "O linexcel não conseguiu concluir esta avaliação: o valor "
            "apresentado é uma alternativa protegida contra erros."
        ),
        "value_volatile": "Não recalculado (volátil)",
        "value_volatile_desc": (
            "TODAY, NOW e RAND respondem de forma diferente a cada cálculo, pelo que "
            "o linexcel mantém o que o ficheiro guarda em vez de inventar um "
            "valor que o livro nunca teve."
        ),
        "external_books": "Livros externos",
        "external_read_folder": "lido da pasta indicada",
        "external_read_cache": "não lido — valor em cache neste ficheiro",
        "external_read_none": "não lido",
        "value_external": "Lido de outro livro",
        "value_external_desc": (
            "Este valor vem de um livro para o qual o linexcel foi apontado, "
            "não do ficheiro analisado: a referência foi substituída pelo valor "
            "antes de a fórmula ser avaliada."
        ),
        "external_hint": (
            "Indique a pasta com estes ficheiros (--refs-dir) para que o "
            "linexcel os possa ler."
        ),
        "query_source": "Origem M",
        "query_loaded": "Carregada em",
        "query_not_loaded": "Não carregada numa folha (apenas ligação)",
        "query_reads": "Lê",
        "query_reads_hint": (
            "Uma origem fora deste livro é nomeada pela consulta, não lida "
            "pelo linexcel: o seu conteúdo não está no ficheiro."
        ),
        "value_col_file": "Ficheiro Excel",
        "value_col_calc": "Recálculo linexcel",
        "value_col_cell": "Célula",
        "value_not_in_file": "Não guardado",
        "value_not_recalc": "Não recalculado",
        "value_no_cache_desc": (
            "O ficheiro não guarda qualquer valor para comparar: o Excel só o "
            "escreve ao gravar um livro calculado."
        ),
        "sampled_cells": "{shown} células amostradas em {count}.",
        "value_match": "O valor recalculado coincide com o ficheiro",
        "value_mismatch": "O valor recalculado difere do ficheiro",
        "target": "Destino",
        "step_decomp": "Decomposição passo a passo",
        "step_hint": "Cada função e operador é avaliado individualmente.",
        "final_result": "Resultado final",
        "not_evaluated": "não avaliado",
        "precedents": "Precedentes",
        "dependents": "Dependentes",
        "cells": "células",
        "ai_doc": "🤖 Documentação de IA (gerada)",
        "ai_vision": "🤖 IA — lido na captura",
        "ai_overview": "🤖 Visão geral gerada por IA",
        "ai_overview_desc": (
            "Esta visão geral foi redigida por um modelo de IA a partir da "
            "linhagem determinista dos cálculos. As informações apresentadas "
            "provêm das fórmulas e dos dados da pasta de trabalho."
        ),
        "fallback": (
            "Não foi possível carregar o Cytoscape (é necessário acesso ao "
            "CDN). O grafo JSON continua disponível através de result.to_dict()."
        ),
        "stats": "{formulas} fórmulas · {nodes} nós · {edges} ligações{vba}",
        "kind_cell": "Fórmula",
        "kind_group": "Fórmulas estendidas",
        "kind_input": "Dados de origem",
        "kind_name": "Célula/intervalo nomeado",
        "kind_vba": "VBA",
        "kind_query": "Power Query",
        "kind_misc": "Outros (agregados)",
        "kind_opaque": "Referência externa",
        "placeholder_title": "Selecione um nó",
        "placeholder_desc": (
            "Selecione um nó no grafo para consultar a sua fórmula, o valor "
            "calculado, a avaliação passo a passo e a documentação gerada por "
            "IA."
        ),
        "sheets_tab": "Folhas",
        "all_sheets": "Todas as folhas",
        "sheet_filter": "Filtro por folha",
        "sheet_dims": "{rows} linhas × {cols} colunas",
        "visibility": "Visibilidade",
        "freeze_panes": "Painéis fixos",
        "hidden_columns": "Colunas ocultas",
        "merged_ranges": "Intervalos unidos",
        "comments": "Comentários",
        "sheet_preview": "Pré-visualização de células",
        "sheet_render": "Folha renderizada",
        "details_panel": "Detalhes do nó",
        "page": "Página {n}",
    },
    "nl": {
        "graph": "Graaf",
        "overview": "Werkmapoverzicht",
        "visual": "Visuele weergave",
        "search": "Zoeken…",
        "search_label": "Cellen en formules doorzoeken (Enter)",
        "search_matches": "Resultaten: {count}",
        "search_none": "Geen resultaten",
        "search_clear": "Zoekopdracht wissen",
        "theme_dark": "Donker thema",
        "fit_all": "Passend maken",
        "fit_sel": "Selectie passend maken",
        "zoom_in": "Inzoomen",
        "zoom_out": "Uitzoomen",
        "layout_flow": "Stroom",
        "layout_organic": "Organisch",
        "close": "Sluiten",
        "formula": "Formule",
        "stretched_pattern": "Doorgetrokken patroon over {count} cellen ({bbox}).",
        "value_heading": "Waarde",
        "value_samples": "Voorbeeldwaarden",
        "value_from_file": "Gelezen uit het Excel-bestand",
        "value_recalc": "Herberekend door linexcel",
        "value_fallback": "terugvalwaarde van linexcel (foutbeveiligd)",
        "value_recalc_desc": (
            "linexcel heeft deze waarde opnieuw berekend op basis van de "
            "gegevens in de werkmap."
        ),
        "value_from_file_desc": (
            "Waarde zoals Excel die in het bestand heeft opgeslagen; linexcel "
            "heeft haar niet herberekend."
        ),
        "value_fallback_desc": (
            "linexcel kon deze evaluatie niet voltooien: de getoonde waarde is "
            "een foutbeveiligde terugvalwaarde."
        ),
        "value_volatile": "Niet herberekend (vluchtig)",
        "value_volatile_desc": (
            "TODAY, NOW en RAND antwoorden bij elke berekening anders, dus linexcel "
            "houdt aan wat het bestand bewaart in plaats van een waarde te "
            "verzinnen die de werkmap nooit had."
        ),
        "external_books": "Externe werkmappen",
        "external_read_folder": "gelezen uit de opgegeven map",
        "external_read_cache": ("niet gelezen — waarde in dit bestand bewaard"),
        "external_read_none": "niet gelezen",
        "value_external": "Uit een andere werkmap gelezen",
        "value_external_desc": (
            "Deze waarde komt uit een werkmap die linexcel kreeg aangewezen, "
            "niet uit het geanalyseerde bestand: de verwijzing werd door de "
            "waarde vervangen voordat de formule werd berekend."
        ),
        "external_hint": (
            "Geef de map met deze bestanden op (--refs-dir) zodat linexcel ze "
            "kan lezen."
        ),
        "query_source": "M-bron",
        "query_loaded": "Geladen in",
        "query_not_loaded": "Niet in een blad geladen (alleen verbinding)",
        "query_reads": "Leest",
        "query_reads_hint": (
            "Een bron buiten deze werkmap wordt door de query genoemd, niet "
            "door linexcel gelezen: de inhoud staat niet in het bestand."
        ),
        "value_col_file": "Excel-bestand",
        "value_col_calc": "Herberekening linexcel",
        "value_col_cell": "Cel",
        "value_not_in_file": "Niet opgeslagen",
        "value_not_recalc": "Niet herberekend",
        "value_no_cache_desc": (
            "Het bestand bevat geen waarde om mee te vergelijken: Excel schrijft "
            "die alleen bij het opslaan van een berekende werkmap."
        ),
        "sampled_cells": "{shown} van {count} cellen bemonsterd.",
        "value_match": "De herberekende waarde komt overeen met het bestand",
        "value_mismatch": "De herberekende waarde wijkt af van het bestand",
        "target": "Doel",
        "step_decomp": "Stapsgewijze ontleding",
        "step_hint": "Elke functie en operator wordt afzonderlijk geëvalueerd.",
        "final_result": "Eindresultaat",
        "not_evaluated": "niet geëvalueerd",
        "precedents": "Voorgangers",
        "dependents": "Afhankelijken",
        "cells": "cellen",
        "ai_doc": "🤖 AI-documentatie (gegenereerd)",
        "ai_vision": "🤖 AI — van de schermafbeelding",
        "ai_overview": "🤖 Door AI gegenereerd overzicht",
        "ai_overview_desc": (
            "Dit overzicht is geschreven door een AI-model op basis van de "
            "deterministische herkomst van de berekeningen. De weergegeven "
            "gegevens komen uit de formules en gegevens van de werkmap."
        ),
        "fallback": (
            "Cytoscape kon niet worden geladen (CDN-toegang vereist). De "
            "JSON-graaf blijft beschikbaar via result.to_dict()."
        ),
        "stats": "{formulas} formules · {nodes} knopen · {edges} verbindingen{vba}",
        "kind_cell": "Formule",
        "kind_group": "Doorgetrokken formules",
        "kind_input": "Brongegevens",
        "kind_name": "Benoemde cel/bereik",
        "kind_vba": "VBA",
        "kind_query": "Power Query",
        "kind_misc": "Overig (samengevoegd)",
        "kind_opaque": "Externe verwijzing",
        "placeholder_title": "Selecteer een knoop",
        "placeholder_desc": (
            "Selecteer een knoop in de graaf om de formule, de berekende "
            "waarde, de stapsgewijze evaluatie en de door AI gegenereerde "
            "documentatie te bekijken."
        ),
        "sheets_tab": "Bladen",
        "all_sheets": "Alle bladen",
        "sheet_filter": "Bladfilter",
        "sheet_dims": "{rows} rijen × {cols} kolommen",
        "visibility": "Zichtbaarheid",
        "freeze_panes": "Geblokkeerde titels",
        "hidden_columns": "Verborgen kolommen",
        "merged_ranges": "Samengevoegde bereiken",
        "comments": "Opmerkingen",
        "sheet_preview": "Celvoorbeeld",
        "sheet_render": "Weergegeven blad",
        "details_panel": "Knoopdetails",
        "page": "Pagina {n}",
    },
    "ja": {
        "graph": "グラフ",
        "overview": "ブック概要",
        "visual": "ビジュアルプレビュー",
        "search": "検索…",
        "search_label": "セルと数式を検索（Enter キー）",
        "search_matches": "{count} 件",
        "search_none": "該当なし",
        "search_clear": "検索をクリア",
        "theme_dark": "ダークテーマ",
        "fit_all": "全体表示",
        "fit_sel": "選択範囲を表示",
        "zoom_in": "拡大",
        "zoom_out": "縮小",
        "layout_flow": "フロー",
        "layout_organic": "オーガニック",
        "close": "閉じる",
        "formula": "数式",
        "stretched_pattern": "{count} セル（{bbox}）にコピーされた数式パターンです。",
        "value_heading": "値",
        "value_samples": "値のサンプル",
        "value_from_file": "Excel ファイルから読み取り",
        "value_recalc": "linexcel が再計算",
        "value_fallback": "linexcel の代替値（エラー保護）",
        "value_recalc_desc": ("linexcel がブックの入力値からこの値を再計算しました。"),
        "value_from_file_desc": (
            "Excel がファイルに保存した値です。linexcel は再計算していません。"
        ),
        "value_fallback_desc": (
            "linexcel はこの評価を完了できませんでした。"
            "表示されている値はエラー保護のための代替値です。"
        ),
        "value_volatile": "再計算なし（揮発性）",
        "value_volatile_desc": (
            "TODAY、NOW、RAND は計算のたびに異なる値を返すため、linexcel は"
            "ブックが持たなかった値を作らず、ファイルの値をそのまま示します。"
        ),
        "external_books": "外部ブック",
        "external_read_folder": "指定フォルダーから読み取り",
        "external_read_cache": "未読み取り — このファイル内のキャッシュ値",
        "external_read_none": "未読み取り",
        "value_external": "他のブックから読み取り",
        "value_external_desc": (
            "この値は linexcel に指定された別のブックのものであり、解析対象の"
            "ファイルのものではありません。数式の評価前に、参照そのものが値に"
            "置き換えられています。"
        ),
        "external_hint": (
            "これらのファイルを含むフォルダーを --refs-dir で指定すると、"
            "linexcel が読み取れます。"
        ),
        "query_source": "M ソース",
        "query_loaded": "読み込み先",
        "query_not_loaded": "シートに読み込まれていません（接続のみ）",
        "query_reads": "参照元",
        "query_reads_hint": (
            "このブックの外にあるソースは、クエリが名前を挙げているだけで、"
            "linexcel は読み取っていません。その内容はファイルにありません。"
        ),
        "value_col_file": "Excel ファイル",
        "value_col_calc": "linexcel 再計算",
        "value_col_cell": "セル",
        "value_not_in_file": "未保存",
        "value_not_recalc": "再計算なし",
        "value_no_cache_desc": (
            "比較できる値がファイルに保存されていません"
            "（Excel は計算済みブックを保存したときにのみ書き込みます）。"
        ),
        "sampled_cells": "{count} セル中 {shown} セルを抽出。",
        "value_match": "再計算値はファイルの値と一致します",
        "value_mismatch": "再計算値はファイルの値と異なります",
        "target": "対象範囲",
        "step_decomp": "ステップごとの分解",
        "step_hint": "各関数・演算子を個別に評価しています。",
        "final_result": "最終結果",
        "not_evaluated": "未評価",
        "precedents": "参照元",
        "dependents": "参照先",
        "cells": "セル",
        "ai_doc": "🤖 AI ドキュメント（生成）",
        "ai_vision": "🤖 AI — スクリーンショットから",
        "ai_overview": "🤖 AI が生成した概要",
        "ai_overview_desc": (
            "この概要は、決定論的に抽出された計算系統に基づいて AI モデルが"
            "作成したものです。記載内容はブックの数式とデータに由来します。"
        ),
        "fallback": (
            "Cytoscape を読み込めませんでした（CDN へのアクセスが必要です）。"
            "JSON グラフは result.to_dict() から引き続き利用できます。"
        ),
        "stats": "数式 {formulas} 件 · ノード {nodes} 件 · エッジ {edges} 件{vba}",
        "kind_cell": "数式",
        "kind_group": "コピーされた数式",
        "kind_input": "ソースデータ",
        "kind_name": "名前付きセル/範囲",
        "kind_vba": "VBA",
        "kind_query": "Power Query",
        "kind_misc": "その他（集約）",
        "kind_opaque": "外部参照",
        "placeholder_title": "ノードを選択してください",
        "placeholder_desc": (
            "グラフ内のノードを選択すると、数式、計算結果、ステップごとの"
            "評価、AI が生成したドキュメントを確認できます。"
        ),
        "sheets_tab": "シート",
        "all_sheets": "すべてのシート",
        "sheet_filter": "シートで絞り込み",
        "sheet_dims": "{rows} 行 × {cols} 列",
        "visibility": "表示状態",
        "freeze_panes": "ウィンドウ枠の固定",
        "hidden_columns": "非表示の列",
        "merged_ranges": "結合されたセル範囲",
        "comments": "コメント",
        "sheet_preview": "セルのプレビュー",
        "sheet_render": "シートの描画",
        "details_panel": "ノードの詳細",
        "page": "ページ {n}",
    },
    "zh": {
        "graph": "图谱",
        "overview": "工作簿概览",
        "visual": "可视化预览",
        "search": "搜索…",
        "search_label": "搜索单元格与公式（回车）",
        "search_matches": "{count} 个匹配",
        "search_none": "无匹配结果",
        "search_clear": "清除搜索",
        "theme_dark": "深色主题",
        "fit_all": "适应窗口",
        "fit_sel": "适应所选",
        "zoom_in": "放大",
        "zoom_out": "缩小",
        "layout_flow": "流向",
        "layout_organic": "有机",
        "close": "关闭",
        "formula": "公式",
        "stretched_pattern": "填充公式模式，覆盖 {count} 个单元格（{bbox}）。",
        "value_heading": "值",
        "value_samples": "取值示例",
        "value_from_file": "读取自 Excel 文件",
        "value_recalc": "由 linexcel 重新计算",
        "value_fallback": "linexcel 回退值（已防错）",
        "value_recalc_desc": ("linexcel 已根据工作簿中的输入重新计算该值。"),
        "value_from_file_desc": (
            "此值由 Excel 保存在文件中，linexcel 未对其重新计算。"
        ),
        "value_fallback_desc": (
            "linexcel 未能完成本次求值：所显示的值是一个已防错的回退值。"
        ),
        "value_volatile": "未重算（易失性）",
        "value_volatile_desc": (
            "TODAY、NOW、RAND 每次计算的结果都不同，"
            "因此 linexcel 保留文件中保存的值，而不是编造工作簿从未有过的值。"
        ),
        "external_books": "外部工作簿",
        "external_read_folder": "已从所提供的文件夹读取",
        "external_read_cache": "未读取 — 使用本文件中缓存的值",
        "external_read_none": "未读取",
        "value_external": "读取自其他工作簿",
        "value_external_desc": (
            "该值来自指定给 linexcel 的另一个工作簿，而非被分析的文件："
            "公式求值前，引用本身已被替换为该值。"
        ),
        "external_hint": (
            "使用 --refs-dir 指定包含这些文件的文件夹，linexcel 即可读取。"
        ),
        "query_source": "M 源代码",
        "query_loaded": "加载到",
        "query_not_loaded": "未加载到工作表（仅连接）",
        "query_reads": "读取",
        "query_reads_hint": (
            "查询指明了本工作簿之外的数据源，但 linexcel 并未读取：其内容不在本文件中。"
        ),
        "value_col_file": "Excel 文件",
        "value_col_calc": "linexcel 重算",
        "value_col_cell": "单元格",
        "value_not_in_file": "文件中无",
        "value_not_recalc": "未重算",
        "value_no_cache_desc": (
            "文件中没有可用于比较的值（Excel 仅在保存已计算的工作簿时才写入）。"
        ),
        "sampled_cells": "已抽样 {count} 个单元格中的 {shown} 个。",
        "value_match": "重算值与文件中的值一致",
        "value_mismatch": "重算值与文件中的值不一致",
        "target": "目标区域",
        "step_decomp": "逐步分解",
        "step_hint": "每个函数和运算符均单独求值。",
        "final_result": "最终结果",
        "not_evaluated": "未求值",
        "precedents": "引用单元格",
        "dependents": "从属单元格",
        "cells": "个单元格",
        "ai_doc": "🤖 AI 文档（自动生成）",
        "ai_vision": "🤖 AI — 读自截图",
        "ai_overview": "🤖 AI 生成的概览",
        "ai_overview_desc": (
            "本概览由 AI 模型根据确定性的计算血缘生成。"
            "所述内容均来自工作簿的公式与数据。"
        ),
        "fallback": (
            "无法加载 Cytoscape（需要访问 CDN）。"
            "JSON 图谱仍可通过 result.to_dict() 获取。"
        ),
        "stats": "{formulas} 个公式 · {nodes} 个节点 · {edges} 条边{vba}",
        "kind_cell": "公式",
        "kind_group": "填充公式",
        "kind_input": "源数据",
        "kind_name": "命名单元格/区域",
        "kind_vba": "VBA",
        "kind_query": "Power Query",
        "kind_misc": "其他（已聚合）",
        "kind_opaque": "外部引用",
        "placeholder_title": "请选择一个节点",
        "placeholder_desc": (
            "在图谱中选择一个节点，即可查看其公式、计算值、逐步求值过程"
            "以及 AI 生成的文档。"
        ),
        "sheets_tab": "工作表",
        "all_sheets": "所有工作表",
        "sheet_filter": "按工作表筛选",
        "sheet_dims": "{rows} 行 × {cols} 列",
        "visibility": "可见性",
        "freeze_panes": "冻结窗格",
        "hidden_columns": "隐藏列",
        "merged_ranges": "合并区域",
        "comments": "批注",
        "sheet_preview": "单元格预览",
        "sheet_render": "工作表渲染图",
        "details_panel": "节点详情",
        "page": "第 {n} 页",
    },
}


def validate_language(language: str) -> str:
    """Return ``language`` if it is supported, else raise ``ValueError``.

    >>> validate_language("ja")
    'ja'
    >>> validate_language("klingon")
    Traceback (most recent call last):
        ...
    ValueError: Unsupported language: 'klingon'. Use one of
    ('en', 'fr', 'es', 'de', 'it', 'pt', 'nl', 'ja', 'zh')
    """
    if language not in LANGUAGES:
        raise ValueError(f"Unsupported language: {language!r}. Use one of {LANGUAGES}")
    return language


def ui_payload(language: str) -> dict[str, Any]:
    """Interface strings to embed for ``language``.

    Only the requested language and the English fallback are shipped: the
    viewer resolves a missing key through ``I18N.en``, so embedding the other
    seven locales in every report would be dead weight.
    """
    validate_language(language)
    payload: dict[str, Any] = {DEFAULT_LANGUAGE: UI_STRINGS[DEFAULT_LANGUAGE]}
    payload[language] = UI_STRINGS[language]
    return payload
