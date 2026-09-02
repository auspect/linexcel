# REFACTOR_PLAN — Repenser `src/linexcel/analyzer.py` en orchestrateur mince

> Proposition d'architecture et plan d'implémentation. Aucun code n'a été écrit.
> Cible : `analyzer.py` (1882 lignes) simplifié et repensé — pas une extraction de plus.
> Baseline vérifiée : branche `main`, commit `f32e4dc`, `uv run pytest tests/ -q` → **586 passed** (~30 s).
> Philippe accepte que les **tests changent** ; la propreté prime sur la rétro-compat privée.

---

## 1. Analyse profonde de l'existant

### 1.1 Carte physique de `analyzer.py` (par plages de lignes)

| Lignes | Contenu | Responsabilité | ~taille |
|---|---|---|---|
| 1–131 | docstring de module, imports, **constantes** (`SCAN_*`, `MAX_*`, budgets, `VOLATILE_FUNCTIONS`, regex volatile) | réglages globaux du pipeline | 131 |
| 133–150 | `FormulaGroup` (dataclass) | un motif de formule R1C1 étiré = futur nœud | 18 |
| 153–212 | `_Budget` | budget (appels + temps) de la décomposition par étapes | 60 |
| 215–739 | `_ValueResolver` | lecture de la valeur d'une cellule (le vrai cœur) | **525** |
| 742–886 | helpers module : `_external_warning`, `_query_warning`, `_external_name`, `_a1_position`, `_as_literal`, `_serial_of`, `_is_volatile`, `a1`, `SECONDS_PER_SHEET_MB`, `WORTH_MENTIONING_SECONDS`, `_SHEET_PART_RE`, `sheet_bytes` | formatage messages / estimation durée / poids zippé | 145 |
| 888–935 | `inspect_workbook` | pré-analyse « combien de temps / que va-t-on laisser de côté » | 48 |
| 938–1597 | **`analyze_workbook`** | le pipeline entier, ~660 lignes, avec 8 fermetures mutuelles | **660** |
| 1600–1683 | quarantaine : regex `_BRACKETED_RE`/`_SPAN_RE`/`_GUARD_RE`/`_SHEET_QUALIFIER_RE`, `_quarantine_unresolvable`, `_is_unresolvable`, `_ensure_scratch` | faire survivre `evaluate_all` à une référence morte | 84 |
| 1685–1882 | helpers de fin : `_collect_ref_strings`, `_bound_names`, `_merge_rects`, `_bbox_a1`, `_chunk_rows`, `_spread_cells`, `_sample_range_values`, `_resolve_call`, `_resolve_vba_write` | AST→références, découpage chunks, échantillonnage, VBA | ~198 |

**Total ≈ 1882 lignes.** Deux blocs concentrent ~63 % du fichier : `_ValueResolver` (525) et `analyze_workbook` (660).

### 1.2 Les phases de `analyze_workbook()` (le flux réel)

En lisant le corps ligne à ligne, le pipeline se découpe en étapes séquentielles nettes :

1. **Structure** (l. 962–989) — `load_workbook` read-only pour `sheet_dims` (avec `tables._force_dimensions`) et `defined_names` (`tables._collect_defined_names`) ; puis liens externes (`external.read_external_links`, `find_workbooks`, `resolve_books`) ; puis valeurs en cache (`loader.load_cached_values`).
2. **Boot moteur** (l. 991–1031) — `formualizer.from_bytes` → `evaluate_all()` ; si échec : **quarantaine** des cellules irrésolvables (`_quarantine_unresolvable`), reconstruction de l'engine depuis les octets, re-tentative ; `engine_alive` bascule à `False` en dernier recours ; création de la feuille scratch (`_ensure_scratch`).
3. **Tables** (l. 1033–1037) — `tables._build_table_index`.
4. **Budget + résolveur** (l. 1039–1051) — instanciation de `_Budget` et `_ValueResolver` (seul moment où tous les faits lus précédemment sont réunis).
5. **Extraction + regroupement** (l. 1053–1121) — balayage par feuille en chunks (`_chunk_rows`, borné par `MAX_CELLS_PER_SHEET`) via `fsheet.get_formulas`, ré-injection des formules quarantaines, regroupement par `canonical_r1c1` en `FormulaGroup` ; produit `groups`, `formula_count`, `sheet_stats`, `cell_owner`.
6. **Sélection des nœuds** (l. 1123–1163) — tri par feuille, plafond `MAX_NODES_PER_SHEET`, création du nœud `misc:{sheet}`, remplissage `cell_owner[cell] = node_id`, `kept_groups` ordonnés.
7. **Construction nœuds/arêtes** (l. 1165–1409) — cœur en 2 moitiés :
   - **fermetures d'infrastructure** : `ensure_opaque_node`, `ensure_input_node`, `add_edge`, `resolve_rect_edges` (elles capturent `nodes/edges/input_nodes/cell_owner/sheet_dims/resolver/table_index/kept_groups` — c'est ici que naît l'enchevêtrement) ;
   - **nœuds « defined names »** (l. 1284–1311) puis **nœuds de formules** (l. 1317–1409) : parse AST (`fz.parse` + cache `ast_cache`), `_collect_ref_strings`, `stretch_ref`, `_merge_rects`, `resolve_rect_edges`, valeur `resolver.describe`, échantillons `_spread_cells`, **décomposition** (`resolver.preload_steps` + `decompose._decompose`), enrichissement table (`tables._enrich_with_table`).
8. **VBA** (l. 1411–1475) — `vba.extract_vba_modules`/`analyze_vba`, nœuds de procédures, arêtes d'appels (`_resolve_call`), arêtes de lecture/écriture (réutilise `ensure_input_node`, `resolve_rect_edges`, `_resolve_vba_write`).
9. **Power Query** (l. 1477–1542) — `powerquery.read_queries`, nœuds de requêtes, arêtes source/upstream/load (réutilise `ensure_input_node`, `resolve_rect_edges`, `name_nodes`, `tables_by_name`), fermeture `query_source_node`.
10. **Assemblage des warnings** (l. 1544–1554) — recovery, uncomputed, external.
11. **Assemblage du graphe** (l. 1556–1597) — dict `meta` (stats, warnings) + `sheets/nodes/edges` ; retour `{graph, engine, analysisId}`.

> `inspect_workbook` (l. 888) est un pipeline « frère » plus petit (structure + liens externes + `sheet_bytes` + plafonds), qui ne fait **aucun** appel à `_ValueResolver` ni aux nœuds.

### 1.3 Ce que fait réellement `_ValueResolver`

C'est **le** client de lecture de valeur, et il porte trois responsabilités entremêlées (mais cohérentes autour d'une seule question : *« quelle est la valeur de cette cellule, et d'où vient-elle, sous quel budget ? »*) :

- **Lecture/récupération** — `value()` est l'entrée unique : moteur → valeur engine → si non-computée : récupération par cellule (`_recover` → `_eval_formula` → branche fallback IFERROR/IFNA via `decompose._guard_fallback_expr`) → sinon valeur du fichier (`_from_cache`). Détection volatile (`_is_volatile` → label `"volatile"`), memo `_resolved`, caches `_compared`/`_uncomputed`.
- **Résolution de chaîne** — `_resolve_precedents` / `_resolve_chain` : marche récursive profondeur ≤ `MAX_CHAIN_DEPTH`, largeur ≤ `MAX_CHAIN_RANGE_CELLS`, puis `_remember` qui réinjecte la valeur récupérée dans l'engine (`set_value`) pour que la suite puisse lire une constante.
- **Passerelle scratch + budget** — `eval_expr` / `_eval_raw` / `preload_steps` (+ `_step_cache`) : c'est la **seule porte** que `decompose` utilise pour évaluer une étape ; c'est aussi le seul endroit où le budget temps/appels (`_Budget`) est honoré. `preload_steps` gèle la mise en cache quand `engine_alive` est faux (l'engine mute pendant la décomposition).
- **Clients externes** — `external_books`, `external_value`, `external_workbooks`, `substitute_externals`, `_book_for`, `_named_books`, `_declared_by_name` : lire d'autres classeurs (dossier ou cache) et substituer la valeur dans une expression avant évaluation.
- **Comptabilité / avertissements** — `describe`, `cached_value`, `_date_text`, `_check_mismatch` (décompte + plafond `MAX_VALUE_WARNINGS`), `uncomputed_warning`.

**Co hésion réelle : « une cellule, une valeur, une source, sous budget ».** Ce n'est pas un ramassis : c'est un objet d'état avec de vraies invariantes (l'engine « mort » qui mute, le cache d'étapes gelé, le budget partagé). Le déplacer *en bloc* dans un module dédié est plus honnête que de le fendre artificiellement — la vraie dette n'est pas dans cette classe, elle est dans le monolithe `analyze_workbook` qui l'instancie et dans le **façade d'import** d'`analyzer.py`.

### 1.4 Les helpers de graphe / nœuds / arêtes

Ils sont **entièrement** à l'intérieur de `analyze_workbook` sous forme de **fermetures mutuelles** : `ensure_opaque_node`, `ensure_input_node`, `add_edge`, `resolve_rect_edges` (l. 1168–1282), plus `query_source_node` (l. 1491). Chaque fermeture capture plusieurs des 10+ dicts partagés (`nodes`, `edges`, `input_nodes`, `cell_owner`, `sheet_dims`, `kept_groups`, `name_nodes`, `table_index`, `resolver`, `defined_names`). C'est **la cause profonde de la taille** : impossible de raisonner phase par phase, tout s'appelle mutuellement à travers des états globaux de fonction.

Les helpers purs de fin de fichier (`_collect_ref_strings`, `_bound_names`, `_merge_rects`, `_bbox_a1`, `_spread_cells`, `_sample_range_values`, `_resolve_call`, `_resolve_vba_write`) sont, eux, **déjà découplés** (pas de capture, paramètres passés) — déplaçables presque à l'identique.

### 1.5 Constantes et ce qu'elles pilotent

| Constante | Piloté par / utilisée dans |
|---|---|
| `SCAN_CHUNK_ROWS`, `SCAN_CHUNK_CELLS`, `_chunk_rows` | bornage des appels `get_formulas` (extraction **et** quarantaine) |
| `SMALL_RANGE_CELLS` | seuil petit/grand rectangle dans `resolve_rect_edges` et `_resolve_vba_write` |
| `MAX_NODES_PER_SHEET` | plafond nœuds/sélection + nœud `misc` + affiché dans `inspect_workbook` |
| `MAX_SCRATCH_EVALS`, `DEFAULT_STEP_SECONDS`, `_Budget` | coût de la décomposition par étapes |
| `MAX_VALUE_SAMPLE` | échantillonnage valeurs (`_sample_range_values`, `_spread_cells`) |
| `MAX_VBA_CODE_CHARS`, `MAX_QUERY_CODE_CHARS`, `MAX_QUERY_SOURCES_SHOWN` | troncature code VBA/PQ, listing sources |
| `MAX_VALUE_WARNINGS`, `MAX_UNCOMPUTED_LISTED` | plafond des avertissements de divergence / non-computées |
| `MAX_CHAIN_DEPTH`, `MAX_CHAIN_RANGE_CELLS` | profondeur/largeur de la marche de récupération |
| `VOLATILE_FUNCTIONS` + regex | détection volatile |
| `SECONDS_PER_SHEET_MB`, `WORTH_MENTIONING_SECONDS`, `_SHEET_PART_RE` | estimation de durée (`sheet_bytes`, `inspect_workbook`, CLI) |

### 1.6 Couplages constatés (synthèse)

1. **`analyze_workbook` est un monolithe séquentiel où 4 phases (sélection, nœuds formules, VBA, Power Query) partagent 10+ dicts via des fermetures.** C'est la dette principale ; la répartition phase par phase est pourtant *quasi-linéaire* (rien de récursif entre phases), donc **découpable**.
2. **`analyzer.py` est aussi un « hub de ré-export »** : il importe `_render_expr`, `SCRATCH_SENTINEL` de `decompose` (avec `# noqa: F401 (re-exported: tests import it from analyzer)`), `CachedValues`, `load_cached_values`, `declared_cells`, `MAX_CELLS_PER_SHEET`, `MAX_DENSE_CELLS`, `_stepped` de `loader`, `serial_to_date_text` de `values`, etc. Les tests et la CLI s'appuient sur ce façade (`from linexcel.analyzer import …`), ce qui *pérennise* le blob. **Ce façade doit être dissous**, pas entretenu.
3. **`_ValueResolver` est une classe cohérente mais volumineuse**, consommée par `decompose` (interface minimale `eval_expr`, `external_value`, `value`, `preload_steps`) — une couture saine à préserver.
4. **Double emploi réel** : `resolve_rect_edges` et `_resolve_vba_write` sont deux parcours quasi identiques cellule→arête (même clippage, même seuil `SMALL_RANGE_CELLS`, même logique petit/grand/approx). ≈40 lignes **unifiables** en un seul helper paramétré par la direction de l'arête.
5. **Cycle potentiel** : `decompose` n'importe `_ValueResolver` que sous `TYPE_CHECKING` (annotation) précisément pour éviter un cycle — signe que la classe doit **quitter** `analyzer.py` vers un module feuille (`resolver.py`), sinon tout module métier qui touche aux valeurs resterait lié à l'orchestrateur.

---

## 2. Architecture cible proposée

### 2.1 Principe

Remplacer le monolithe impératif + ses fermetures mutuelles par une **orchestration séquentielle explicite** : `analyzer.py` appelle des **phases monométier**, chacune recevant un **contexte de lecture** (faits du classeur) et rendant un **résultat intermédiaire dataclass**, sans état mutuel partagé entre modules. Les seuls états qui *accumulent* (nœuds, arêtes, cell_owner, name_nodes) sont **possédés par une classe `GraphBuilder`** — unique responsabilité « produire les nœuds/arêtes ».

### 2.2 Découpage proposé en modules (ajusté au couplage réel)

```
src/linexcel/
├── analyzer.py    # ORCHESTRATEUR MINCE (réécrit, ~100-140 lignes)
│                  #   analyze_workbook() : appelle les phases dans l'ordre + assemble meta
│                  #   inspect_workbook()  : ré-exporté depuis structure (API publique conservée)
├── engine.py      # [nouveau] vie du moteur formualizer + quarantaine
├── sweep.py       # [nouveau] extraction formules + regroupement R1C1
├── graph.py       # [nouveau] sélection des nœuds + construction nœuds/arêtes (dont names/VBA/PQ)
├── resolver.py    # [nouveau] _Budget + _ValueResolver + helpers valeur/externe/volatile
├── structure.py   # [nouveau] inspect_workbook + sheet_bytes + estimation durée + lecture structure
└── (déjà extraits) decompose.py, tables.py, loader.py, values.py, refs.py, external.py,
                    powerquery.py, vba.py, progress.py, rewrite.py, result.py, insights.py …
```

Chaque module garde **une seule responsabilité** ; le rôle et les dépendances de chacun ci-dessous.

---

### 2.3 Fiches modules cibles

#### `engine.py` — « faire vivre et faire survivre le moteur » (~150 lignes)
- **Responsabilité** : instancier l'engine depuis les octets, lancer `evaluate_all` (avec re-tentative après quarantaine), ajouter la feuille scratch.
- **Contenu** : `boot_engine(data, sheet_dims, warnings) -> EngineSession` ; dataclass `EngineSession(engine, engine_sheets, engine_alive, quarantined, scratch_ready)` ; quarantaine `_quarantine_unresolvable`, `_is_unresolvable`, regex `_BRACKETED_RE`/`_SPAN_RE`/`_GUARD_RE`/`_SHEET_QUALIFIER_RE` ; `_ensure_scratch`. Constantes de balayage `SCAN_CHUNK_ROWS`, `SCAN_CHUNK_CELLS` + `_chunk_rows` (utilisées par quarantaine **et** extraction).
- **Dépendances** : `formualizer`, `progress.Reporter`, `tables` (rien) — n'a pas besoin du résolveur ni des nœuds.

#### `sweep.py` — « extraire les formules et les regrouper en motifs » (~150 lignes)
- **Responsabilité** : balayer les feuilles en chunks et regrouper par `canonical_r1c1` → motifs R1C1.
- **Contenu** : dataclass `FormulaGroup` (déplacée ici) ; `sweep_sheets(engine, sheet_dims, engine_sheets, quarantined, warnings, reporter, budget_cells) -> SweepResult` ; dataclass `SweepResult(groups: dict[(sheet, r1c1), FormulaGroup], formula_count, sheet_stats)`.
- **Dépendances** : `formualizer`, `rewrite.canonical_r1c1`, `loader.MAX_CELLS_PER_SHEET`, `progress.Reporter`.
- **N'exige ni resolver ni table ni nœud** — unitaire et testable en isolation.

#### `graph.py` — « construire les nœuds et les arêtes du graphe » (~550 lignes, le gros morceau)
- **Responsabilité** : transformer motifs R1C1 + defined names + VBA + Power Query + tables en la liste finale `nodes`/`edges`.
- **Contenu** :
  - classe `GraphBuilder` possédant `nodes`, `edges`, `input_nodes`, `name_nodes`, `cell_owner`, `ast_cache` ;
  - méthodes devenues **publiques** (ex-fermetures) : `add_edge`, `ensure_opaque_node`, `ensure_input_node`, `resolve_rect_edges` ;
  - **sélection des nœuds** (plafond `MAX_NODES_PER_SHEET`, nœud `misc`, `cell_owner`, `kept_groups`) ;
  - nœuds **defined names**, nœuds **formules** (+ décomposition via `decompose`), nœuds **VBA**, nœuds **Power Query** ;
  - helpers purs : `_collect_ref_strings`, `_bound_names`, `_merge_rects`, `_bbox_a1`, `_spread_cells`, `_sample_range_values`, `_resolve_call`, `_resolve_vba_write`.
  - Constantes graph : `MAX_NODES_PER_SHEET`, `SMALL_RANGE_CELLS`, `MAX_VALUE_SAMPLE`.
- **Dépendances** : `formualizer`, `resolver._ValueResolver`, `tables._enrich_with_table`, `decompose`, `vba`, `powerquery`, `external`, `refs`, `rewrite`, `values._jsonable`, `progress`.
- **N'écrit aucune donnée ailleurs que dans ses propres `nodes`/`edges`** ; ne lit rien du classeur (tout est injecté en entrée).

#### `resolver.py` — « la valeur d'une cellule, sous budget » (~520 lignes déplacées)
- **Responsabilité** : unique point de lecture de la valeur d'une cellule (engine → récupération → fallback → cache → externe), avec le budget de décomposition.
- **Contenu** : `_Budget`, `_ValueResolver` (déplacés quasi tels quels) ; constants `MAX_SCRATCH_EVALS`, `DEFAULT_STEP_SECONDS`, `MAX_CHAIN_DEPTH`, `MAX_CHAIN_RANGE_CELLS`, `MAX_VALUE_WARNINGS`, `MAX_UNCOMPUTED_LISTED`, `VOLATILE_FUNCTIONS` + regex ; helpers `_is_volatile`, `_external_name`, `_a1_position`, `_as_literal`, `_serial_of`, `_external_warning` (ré-exportés/utilisés ici).
- **Dépendances** : `formualizer`, `external`, `decompose` (`_guard_fallback_expr`, `SCRATCH_SHEET`, `SCRATCH_SENTINEL`), `loader.CachedValues`, `loader.MAX_*`, `values`, `rewrite.qualify_sheet`, `refs`.
- **Feuille du graphe d'import** : plus aucun module métier n'aura à importer `analyzer` (suppression du `TYPE_CHECKING` vers `analyzer._ValueResolver` dans `decompose`, remplacé par `resolver._ValueResolver`).

#### `structure.py` — « ce que le fichier dit de lui-même, avant d'analyser » (~120 lignes)
- **Responsabilité** : pré-analyse légère + lecture de structure du classeur (dimensions, noms définis).
- **Contenu** : `inspect_workbook`, `sheet_bytes`, `SECONDS_PER_SHEET_MB`, `WORTH_MENTIONING_SECONDS`, `_SHEET_PART_RE` ; fonction `read_structure(data) -> Structure(sheet_dims, defined_names)` regroupant la phase 1 d'`analyze_workbook` (openpyxl read-only + `tables._force_dimensions`/`_collect_defined_names`).
- **Dépendances** : `openpyxl`, `zipfile`, `external.read_external_links`, `loader.declared_cells`/`MAX_CELLS_PER_SHEET`/`MAX_DENSE_CELLS`, `tables`.
- *Alternative acceptable (ponytail)* : si on veut un fichier de moins, plier cette phase dans `loader.py` (qui lit déjà la structure/dimensions). Recommandé en module distinct car la CLI en dépend via `sheet_bytes`/`inspect_workbook` et que `loader` ne doit pas dépendre d'`external`.

#### `analyzer.py` — l'orchestrateur mince (réécrit, ~100–140 lignes)
- **Responsabilité** : ordonner les phases et assembler le dictionnaire de sortie. **Aucune logique métier.**
- **Contenu** :
  1. `read_structure` → 2. `engine.boot_engine` → 3. `tables._build_table_index` → 4. `loader.load_cached_values` → 5. instancier `resolver._Budget` + `_ValueResolver` → 6. `sweep.sweep_sheets` → 7. construire `graph.GraphBuilder`, appeler ses étapes (sélection → names → formules → VBA → PQ) → 8. assembler les warnings → 9. assembler `meta`/`graph` → retour `{graph, engine, analysisId}`.
  - `inspect_workbook` conservé **ré-exporté** (`from linexcel.structure import inspect_workbook`) pour ne pas casser la CLI ; l'implémentation vit dans `structure`.
- **Dépendances** : les modules ci-dessus + `values`/`loader`/`result`-agnostique. Il **reste le seul** fichier à importer les modules de phase (aucun module métier n'importe `analyzer`).

### 2.4 Flux et objets intermédiaires (le cœur du désenchevêtrement)

Pour éliminer les fermetures mutuelles, chaque phase rend un objet et le suivant en consomme :

```
data, filename, refs_dir, step_seconds
   │
   ▼  structure.read_structure(data)              → Structure(sheet_dims, defined_names)
   │  external.read_external_links + resolve     → externals, refs_files
   ▼  loader.load_cached_values(data, ...)       → CachedValues
   ▼  engine.boot_engine(data, sheet_dims, ...)  → EngineSession(engine, engine_sheets, engine_alive, quarantined, scratch_ready)
   ▼  tables._build_table_index(...)             → table_index
   ▼  resolver = _ValueResolver(EngineSession, CachedValues, table-independent…)
   ▼  sweep.sweep_sheets(EngineSession, ...)     → SweepResult(groups, formula_count, sheet_stats)
   ▼  builder = GraphBuilder(resolver, defined_names, table_index, sheet_dims, engine_sheets,
   │                         SweepResult, externals, refs_dir, warnings, reporter)
   │  builder.build_names(); builder.build_formula_nodes()
   │  builder.build_vba(macro_files); builder.build_queries()
   ▼  → nodes, edges (accesseurs), kept_groups, n_recovered…
   ▼  assemble meta + warnings + {graph, engine, analysisId}
```

Dataclasses à introduire pour **porter** les données entre modules (au lieu de dicts mutés par 5 mains) :
- `EngineSession` (engine.py) — encapsule `engine, engine_sheets, engine_alive, quarantined, scratch_ready` aujourd'hui passés un par un à 4 fonctions.
- `SweepResult` (sweep.py) — `groups, formula_count, sheet_stats`.
- `Structure` (structure.py) — `sheet_dims, defined_names`.
- (Optionnel) `AnalysisResult` interne dans `analyzer` pour la sortie, mais le retour public reste `{graph, engine, analysisId}` **inchangé**.

Règle de non-circulation : un module de phase reçoit **seulement** les faits qu'il lit et **rend** ses propres produits ; le seul objet « vivant » qui traverse plusieurs phases est `GraphBuilder` (possesseur légitime de `nodes`/`edges`).

### 2.5 API publique et ré-export : ce qu'on garde, ce qu'on dissout

**API publique à préserver à l'identique** (contrat, pas question de le casser) :
- `linexcel.analyze` / `LineageResult` (résident dans `result.py`, inchangés) ;
- `linexcel.analyze_workbook` (importé par `__init__`, `result`, CLI) — **signature et retour identiques** `{graph, engine, analysisId}` ; l'`engine` est restitué car `LineageResult` le garde pour les lectures ultérieures ;
- `linexcel.analyzer.inspect_workbook` (CLI) — ré-exporté depuis `structure`.

**Façade privé à dissoudre** (le vrai « blob ») — symboles actuellement ré-exportés depuis `analyzer` par simple transit (`# noqa: F401`) et **ré-importés ailleurs** :
`_render_expr`, `SCRATCH_SENTINEL` (vivent dans `decompose`) ; `CachedValues`, `load_cached_values`, `declared_cells`, `MAX_CELLS_PER_SHEET`, `MAX_DENSE_CELLS`, `_stepped` (vivent dans `loader`) ; `serial_to_date_text` (vit dans `values`).
→ **Les points d'import (tests + src) doivent pointer vers le module d'origine**, pas vers `analyzer`. C'est un changement de tests assumé.

**Symboles privés utilisés par les tests**, et leur nouvelle maison (à mettre à jour dans les tests) :

| Symbole importé aujourd'hui de `analyzer` | Maison cible |
|---|---|
| `analyze_workbook`, `inspect_workbook`, `sheet_bytes` | `analyzer` (orchestrateur) / `structure` |
| `_is_unresolvable`, `_chunk_rows`, `SCAN_CHUNK_ROWS`, `SCAN_CHUNK_CELLS` | `engine` |
| `FormulaGroup` | `sweep` |
| `_spread_cells`, `_collect_step_exprs` (venant de decompose) | `graph` / `decompose` |
| `_Budget`, `_ValueResolver`, `_is_volatile`, `SCRATCH_SENTINEL` (via decompose) | `resolver` |
| `MAX_CELLS_PER_SHEET`, `MAX_DENSE_CELLS`, `declared_cells`, `load_cached_values`, `CachedValues` | `loader` (ré-import) |

Les tests qui ne touchent **que** des valeurs/structure (`test_progress`, `test_robustness`, `test_values`, `test_lineage` chunk/quarantaine, `test_cli`) verront leur import déplacé vers `loader`/`structure`/`engine`/`resolver` ; le contenu des assertions reste identique.

---

## 3. Plan d'implémentation incrémental (étapes sûres, ordonnées, testables)

### 3.0 Filet de sécurité avant/après (à faire en premier — indispensable)

Le refactor accepte des **changements de tests**, donc on ne peut pas compter sur « les 586 tests restent verts » comme seule preuve. On construit une **comparaison comportementale de haut niveau** :

1. **Barrière haute (invariante)** : les tests qui passent par l'API publique `analyze()`/`analyze_workbook()` sur des classeurs de test et qui vérifient des **propriétés du graphe JSON** (`nodes[id]`, `edges`, `steps`, `valueSource`, warnings) doivent rester verts **quel que soit** le refactor interne. Ce sont eux le « contrat de comportement » : `test_lineage` (voie publique), `test_values` (via `graph_of`/`analyze_workbook`), `test_external`, `test_powerquery`, `test_robustness`. → **Ne jamais les réécrire pour la commodité du refactor** ; s'ils cassent, c'est une régression réelle.
2. **Snapshot de graphe** : avant tout déplacement, générer, pour un petit jeu de classeurs (`tests/fixtures`), le **JSON canonique complet** (nodes+edges+meta triés et normalisés) via un script jetable `scripts/freeze_graph.py` (hors suite, exécuté à la main). Chaque étape rejoue le script et **diff** : la sortie doit être **identique octet pour octet** entre deux étapes qui ne changent que la *structure interne*.
3. **Exécution de la suite complète** après chaque étape (`uv run pytest tests/ -q`), plus les fichiers ciblés listés par étape.

Règle d'acceptation par étape : **sortie du snapshot inchangée** + **tests publics verts** + **tests ciblés de l'étape verts**. Une étape qui échoue se défaire avant d'aller plus loin (le dépôt est sur `main`, commits petits et réversibles).

### 3.1 Les étapes

> Chaque étape est indépendante, réversible, et se vérifie seule. On **ne touche jamais au comportement** dans une étape de déplacement ; toute simplification comportementale est reléguée à l'étape 7 (et alors verrouillée par snapshot).

**Étape 1 — Dissoudre le façade de ré-export (0 risque de comportement, gros gain de lisibilité).**
- Déplacer les imports privés des tests/src : remplacer `from linexcel.analyzer import (CachedValues, load_cached_values, declared_cells, MAX_CELLS_PER_SHEET, MAX_DENSE_CELLS, _stepped, serial_to_date_text, SCRATCH_SENTINEL, _render_expr, _collect_step_exprs, _is_uncomputed …)` par leurs modules d'origine (`loader`, `values`, `decompose`).
- Fichiers touchés : `tests/test_robustness.py`, `tests/test_values.py`, `tests/test_progress.py`, `tests/test_lineage.py`, `tests/test_cli.py` (les imports) — `analyzer.py` lui-même intouché.
- Risque : faible (pur déplacement d'import). Vérif : `uv run pytest tests/test_values.py tests/test_robustness.py tests/test_progress.py tests/test_cli.py -q` puis suite complète.
- Effet : `analyzer.py` cesse d'être un hub ; on voit qu'il ne reste que du métier.

**Étape 2 — Extraire `structure.py` (inspect_workbook + sheet_bytes + estimation + lecture structure).**
- Créer `structure.py` : y déplacer `sheet_bytes`, `_SHEET_PART_RE`, `SECONDS_PER_SHEET_MB`, `WORTH_MENTIONING_SECONDS`, `inspect_workbook`, et une fonction `read_structure(data) -> Structure(sheet_dims, defined_names)` (le bloc structure de `analyze_workbook` l. 962–974).
- `analyzer.analyze_workbook` appelle `read_structure` ; `analyzer.inspect_workbook` devient un ré-export (`from linexcel.structure import inspect_workbook`). Mettre à jour les imports CLI (`cli.py`) vers `structure` pour `sheet_bytes`/constantes/inspect.
- Fichiers : `structure.py` (nouveau), `analyzer.py`, `cli.py`, imports tests.
- Risque : faible. Vérif : `pytest test_cli.py test_progress.py` (inspect/sheet_bytes) + suite + snapshot.
- Effet : `analyzer.py` ne porte plus aucun code « pré-analyse ».

**Étape 3 — Extraire `engine.py` (boot + quarantaine + scratch).**
- Créer `engine.py` : dataclass `EngineSession` + `boot_engine(data, sheet_dims, warnings)` encapsulant exactement le bloc l. 991–1031 (from_bytes → evaluate_all → quarantaine → re-tentative → `engine_alive` → `_ensure_scratch`) ; déplacer `_quarantine_unresolvable`, `_is_unresolvable`, les 4 regex, `_ensure_scratch`, et les constantes de chunk `SCAN_CHUNK_ROWS`/`SCAN_CHUNK_CELLS`/`_chunk_rows`.
- `analyze_workbook` remplace le bloc par `session = engine.boot_engine(data, sheet_dims, warnings)`.
- Fichiers : `engine.py` (nouveau), `analyzer.py`, imports tests (`_is_unresolvable`, `_chunk_rows`, `SCAN_*`).
- Risque : moyen (bloc avec branche d'erreur délicate) — on déplace **tel quel**, sans réécrire. Vérif : `pytest test_lineage.py -k "chunk or row or unroll or scan"` + `test_robustness.py` + snapshot (branches de quarantaine couvertes par fixtures) + suite.
- Effet : le code moteur sort ; `analyze_workbook` commence à raccourcir.

**Étape 4 — Extraire `resolver.py` (budget + résolveur + helpers valeur/externe/volatile).**
- Créer `resolver.py` : y déplacer (à l'identique) `_Budget`, `_ValueResolver`, et les helpers module qui ne servent qu'au résolveur : `_is_volatile` (+ regex/`VOLATILE_FUNCTIONS`), `_external_name`, `_a1_position`, `_as_literal`, `_serial_of`, `_external_warning` ; les constantes budget/chaîne : `MAX_SCRATCH_EVALS`, `DEFAULT_STEP_SECONDS`, `MAX_CHAIN_DEPTH`, `MAX_CHAIN_RANGE_CELLS`, `MAX_VALUE_WARNINGS`, `MAX_UNCOMPUTED_LISTED`, `MAX_VALUE_SAMPLE` (partagée avec graph — laisser la constante à l'endroit qui la lit, sinon ré-exporter).
- Mettre à jour `analyzer.py` (import de `resolver`) et `decompose.py` (son `TYPE_CHECKING` pointe désormais `resolver._ValueResolver`).
- Fichiers : `resolver.py` (nouveau), `analyzer.py`, `decompose.py`, imports tests (`_Budget`, `_ValueResolver`, `_is_volatile`, `SCRATCH_SENTINEL` reste dans `decompose`).
- Risque : moyen (grosse classe, mais déplacement strict). Vérif : `pytest test_values.py -k "Guarded or Scratch or Resolver or Volatile or Budget"` + `test_robustness.py` (budget) + snapshot + suite.
- Effet : la classe de 525 lignes quitte `analyzer` ; fin du risque de cycle d'import.

**Étape 5 — Extraire `sweep.py` (extraction + regroupement + FormulaGroup).**
- Créer `sweep.py` : déplacer `FormulaGroup` et le bloc d'extraction+regroupement (l. 1053–1121) dans `sweep_sheets(...) -> SweepResult`, qui prend en entrée l'`EngineSession` (donc `quarantined`, `engine_sheets`), `sheet_dims`, warnings, reporter, budget.
- `analyze_workbook` appelle `sweep.sweep_sheets` au lieu du `with reporter.phase(...)` inline.
- Fichiers : `sweep.py` (nouveau), `analyzer.py`, imports tests (`_chunk_rows` déjà parti en engine).
- Risque : faible-moyen. Vérif : `pytest test_lineage.py -k "group or pattern or sweep or scan or stretched"` + snapshot + suite.
- Effet : le balayage I/O sort ; `analyze_workbook` se vide d'une boucle lourde.

**Étape 6 — Extraire `graph.py` : GraphBuilder + tout nœud/arête + VBA + Power Query. (l'étape clé)**
- Créer `graph.py` avec la classe `GraphBuilder`. Convertir les **fermetures** en **méthodes** (elles perdent leur capture implicite et prennent des champs `self.`), puis extraire, dans l'ordre :
  1. la **sélection** (plafond `MAX_NODES_PER_SHEET`, nœud `misc`, `cell_owner`, `kept_groups`) — aujourd'hui l. 1123–1163 ;
  2. `ensure_opaque_node`, `ensure_input_node`, `add_edge`, `resolve_rect_edges` + helpers AST (`_collect_ref_strings`, `_bound_names`, `_merge_rects`, `_bbox_a1`, `_spread_cells`, `_sample_range_values`) ;
  3. **nœuds defined names** (l. 1284–1311) ;
  4. **nœuds formules + étapes** (l. 1317–1409) ;
  5. **VBA** (l. 1411–1475) : `build_vba(vba_modules, vba_procs, addins…)` + `_resolve_call`, `_resolve_vba_write` ;
  6. **Power Query** (l. 1477–1542) : `build_queries(queries)` + fermeture `query_source_node`.
- `analyze_workbook` instancie `GraphBuilder(...)` puis appelle les méthodes dans l'ordre ; il ne manipule plus `nodes`/`edges` directement — il les lit en sortie du builder.
- Fichiers : `graph.py` (nouveau), `analyzer.py`, imports tests (`_spread_cells`, `_collect_step_exprs` reste decompose).
- Risque : **élevé** (c'est le désenchevêtrement central). Mesures : faire de l'étape une suite de micro-commits (un sous-bloc à la fois, chacun rejouant le snapshot), ne **réécrire aucune logique**, uniquement « fermeture → méthode ». Vérif après chaque sous-bloc : snapshot **identique** + `pytest test_lineage.py test_values.py test_external.py test_powerquery.py -q` + suite.
- Effet : `analyze_workbook` perd ~500 lignes ; les nœuds/arêtes sont regroupés dans un module unique et testable.

**Étape 7 — Réécrire `analyze_workbook` en orchestrateur mince + nettoyage.**
- `analyze_workbook` ne contient plus que : le contexte (warnings, reporter, timing), l'appel séquentiel des phases (read_structure → externals → cached → boot → tables → resolver → sweep → builder → assembly), l'assemblage `meta`+`graph` et le retour `{graph, engine, analysisId}`. Suppression de `_v`, des 8 fermetures, de tous les dicts intermédiaires désormais détenus par le builder.
- **Simplifications autorisées maintenant** (verrouillées par snapshot) :
  - **Unifier `resolve_rect_edges` et `_resolve_vba_write`** en un seul parcours rectangle→arêtes paramétré par la direction (≈40 lignes en moins, double emploi réel constaté §1.6) ;
  - retirer tout commentaire `# noqa: F401 re-exported` devenu sans objet ;
  - supprimer la dataclass/boucle `per_sheet_groups` intermédiaire si le tri reste lisible ;
  - consolider les docstrings de `analyze_workbook` (le docstring de module reflète déjà les 7 étapes).
- Fichiers : `analyzer.py` (réécrit), éventuellement un micro-ajustement `graph.py`.
- Risque : moyen (réécriture, mais sur des blocs déjà extraits). Vérif : snapshot **identique** + suite complète ; audit visuel : `wc -l analyzer.py` doit être ~100–140.
- Effet : **livrable principal atteint** — `analyzer.py` est un orchestrate ur mince et lisible.

**Étape 8 — Mise au propre des tests privés + bascule des imports restants + CI.**
- Mettre à jour les quelques tests restants qui importent encore depuis `analyzer` des symboles désormais ailleurs (selon la table §2.5). Reformuler uniquement les imports, **jamais les assertions de comportement**.
- Vérifier qu'aucun module de `src/` n'importe `analyzer` sauf les points publics (`__init__`, `result`, `cli`, et les modules de phase appelés par l'orchestrateur — mais aucun module métier ne doit dépendre de `analyzer`). Idéalement : `from linexcel.analyzer import analyze_workbook, inspect_workbook` seulement.
- Lancer : suite complète, `uv run pytest tests/ -q`, et un dernier diff du snapshot de bout en bout (avant la 1re étape vs maintenant) pour prouver que le comportement public est inchangé.
- Effet : dépôt propre, tests verts, aucun reliquat de façade.

### 3.2 Bilan estimé

| Fichier | Avant | Après (est.) |
|---|---|---|
| `analyzer.py` | 1882 | **~110–150** (orchestrateur + ré-export inspect) |
| `resolver.py` | — | ~520 (déplacé de analyzer) |
| `graph.py` | — | ~500–560 (déplacé de analyzer, −40 après unif.) |
| `engine.py` | — | ~150 |
| `sweep.py` | — | ~150 |
| `structure.py` | — | ~120 |
| Total | — | **similaire en volume global**, mais chaque module ≤ ~560 lignes, monométier, sans états croisés |

Volume global quasi identique (on ne *réduit* pas le code métier — il existe pour de bonnes raisons) ; ce qui change est la **structure** : plus de monolithe de 1882 lignes ni de façade, des frontières par responsabilité, et `analyzer.py` lisible d'un coup d'œil.

---

## 4. Ce qui peut être simplifié / supprimé (opportunités recensées)

**Code mort / superflu :**
- Le **façade de ré-export** (`# noqa: F401 (re-exported: tests import it from analyzer)` sur `_render_expr`, etc.) — suppression pure après bascule des imports (étape 1).
- Les **fermetures mutuelles** de `analyze_workbook` qui n'existent que pour capturer des dicts partagés — remplacées par des méthodes de `GraphBuilder`.

**Complexité inutile / à unifier :**
- `resolve_rect_edges` (l. 1249) et `_resolve_vba_write` (l. 1857) : deux parcours rectangle→arêtes quasi identiques (même clippage, même seuil `SMALL_RANGE_CELLS`, mêmes branches petit/grand/approx) — **≈40 lignes unifiables** en un helper commun paramétré par la direction de l'arête.
- `a1(row,col)` est un pur alias de `num_to_col` + f-string déjà présent dans `refs` ; il peut vivre dans `refs` (avec `num_to_col`) pour centraliser la conversion A1.

**Pièges à ne PAS « simplifier » (à garder tels quels) :**
- Le budget `_Budget` à double plafond (appels *et* temps) et la distinction décrite dans les docstrings — une simplification « chrono seul » a déjà causé un bug documenté ;
- la ré-injection de formules quarantaines pendant le balayage (sinon le graphe est vide sous ces cellules) ;
- la logique `preload_steps`/`_engine_alive` (batching gelé quand l'engine mute) ;
- le retour de l'`engine` dans le payload public (`LineageResult` le consomme).

**Observations utiles pour l'implémenteur :**
- `decompose.py` et `tables.py` sont déjà bien séparés ; leur API interne est saine — on **réutilise** `_stepped`, `_chunk_rows` (déplacés), `_render_expr`, etc. plutôt que de réinventer.
- Le graphe d'import cible est **acyclique et dirigé** : `structure/engine/sweep` (feuilles) → `tables/resolver/loader/values/decompose` → `graph` → `analyzer` (racine). Aucun module bas ne doit importer `analyzer`.

---

## 5. Conclusion

La dette n'est pas « un gros fichier à réduire » : c'est **un orchestrateur qui déborde de logique métier + un façade d'import qui le cimente**. La refonte remet chaque responsabilité dans un module dédié (`engine`, `sweep`, `graph`, `resolver`, `structure`) et réduit `analyzer.py` à un **orchestrateur mince** qui ne fait qu'ordonner des phases. Le refactor est découpé en **8 étapes sûres et réversibles**, chacune protégée par un **snapshot de graphe identique** + une **barrière haute de tests publics** qui ne doivent jamais être réécrits — garantissant que, malgré des changements de tests privés assumés, le comportement observé (graphe, warnings, valeurs, étapes) est strictement préservé.
