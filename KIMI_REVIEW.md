# REVUE CROISÉE CRITIQUE — Refactor `analyzer.py` → orchestrateur mince

**Date :** 2026-09-02 · **Revue de :** agent indépendant (2e regard)
**Cible relue :** `REFACTOR_PLAN.md` (337 l.), `src/linexcel/analyzer.py` (1882 l., branch `refactor/analyzer-orchestrator`, commit `e06f269` = plan seul, code intact), `decompose.py`, `tables.py`, `loader.py`, et les tests `test_lineage.py`, `test_values.py`, `test_robustness.py`, `test_progress.py`, `test_powerquery.py`, `test_cli.py`.
**Périmètre :** aucune modification de code. Ce document est un audit, pas une implémentation.

---

## 1. Solidité du plan (découpage en 8 étapes, frontières de modules)

### 1.1 Verdict global
Le plan est **bon dans sa vision** et remarquablement fidèle au code : la carte physique (§1.1), la lecture des phases (§1.2) et la fiche `_ValueResolver` (§1.3) sont exactes. L'architecture cible (engine/sweep/graph/resolver/structure, avec `GraphBuilder` comme seul possesseur de l'état cumulé `nodes/edges/cell_owner`) est cohérente avec le **couplage réel** : les 4 fermetures `ensure_opaque_node` (l.1168), `ensure_input_node` (l.1202), `add_edge` (l.1233), `resolve_rect_edges` (l.1249) + `query_source_node` (l.1491) capturent bien les mêmes ~10 dicts, et le flux est effectivement **quasi-linéaire** (pas de récursion entre phases). Le découpage en modules est le bon.

Deux ordres de problèmes :

**(a) Le plan sous-estime le coût réel de l'étape 1** (« 0 risque », « pur déplacement d'import »). C'est faux et c'est l'étape qui va casser le plus de choses en premier — parce que le façade n'est pas seulement consommé par des `import` en tête de fichiers de test, mais par **~11 sites runtime** (`monkeypatch.setattr(analyzer, …)`, `mock.patch("linexcel.analyzer.…")`, et des `from linexcel.analyzer import …` **à l'intérieur des corps de test**). Détail section 2.3.

**(b) Le siting de certaines constantes crée des dépendances ascendantes non vues** (section 2.4), et l'« unification » de l'étape 7 concerne deux fonctions qui ne sont **pas** aussi identiques que le plan le prétend (section 2.2).

### 1.2 Frontières engine/sweep — un couplage à réconcilier dans les fiches
`_chunk_rows` (l.1784) lit `SCAN_CHUNK_ROWS`/`SCAN_CHUNK_CELLS` et est utilisé **à la fois** par la quarantaine (l.1637, `engine`) **et** par la boucle d'extraction du balayage (l.1068, `sweep`). Le plan envoie `_chunk_rows` + les deux `SCAN_*` en `engine` (étape 3), puis crée `sweep.py` (étape 5) sans le noter en dépendance : la fiche `sweep` (§2.3) liste `rewrite.canonical_r1c1`, `loader.MAX_CELLS_PER_SHEET`, `progress.Reporter` mais **pas `engine._chunk_rows`**. Il faut que `sweep` importe `_chunk_rows` depuis `engine` (feuille → OK, pas de cycle), et le lister. Mineur mais à corriger pour que la vérif de l'étape 5 soit cohérente.

### 1.3 Étape 7 — unification `resolve_rect_edges` / `_resolve_vba_write` : attention, elles ne sont PAS isomorphes
Le plan (§1.6, §4, étape 7) les présente comme « ≈40 lignes unifiables, même clippage, même seuil, mêmes branches petit/grand/approx ». En relisant :
- `resolve_rect_edges` (l.1249-1282) a **trois** branches : petit-range (itération cellule→owner), **grand-range approx sur `kept_groups`/bbox** (l.1274-1282), et cas « feuille absente ».
- `_resolve_vba_write` (l.1857-1882) n'a **pas** de branche « grand-range approx » : sa branche `else` (l.1879-1880) met seulement `has_plain = True` puis ajoute un `ensure_input_node(clipped)` (l.1882). Elle ne parcourt jamais `kept_groups`.
Unifier naïvement « par la direction de l'arête » risque de **faire entrer l'approx-à-la-formule dans les VBA-write** (changerait le graphe) ou de **faire perdre l'approx des formules** (aussi un changement). L'unification est faisable mais doit être paramétrée par un flag « approx-grand-range autorisé », pas seulement par la direction. La fiche §4 « à simplifier » ne le dit pas → risque réel pour le snapshot à l'étape 7.

### 1.4 Constante `MAX_NODES_PER_SHEET` → dépendance upward `structure → graph` non vue
`inspect_workbook` expose `ceilings.nodesPerSheet: MAX_NODES_PER_SHEET` (l.932). Le plan met `inspect_workbook` + `read_structure` dans **`structure`** (feuille, qui ne doit dépendre ni de graph ni d'analyzer), mais **place `MAX_NODES_PER_SHEET` dans `graph`** (§2.3 fiche graph : « Constantes graph : MAX_NODES_PER_SHEET, SMALL_RANGE_CELLS, MAX_VALUE_SAMPLE »). Si `inspect_workbook` lit `graph.MAX_NODES_PER_SHEET`, on obtient `structure → graph`, c'est-à-dire **une flèche vers le haut** qui casse le graphe d'import cible « acyclique et dirigé » (§2.5, §4). Correction : la constante doit vivre dans **`structure`** (avec `inspect_workbook`) ou dans `loader`, et être **importée** par `graph`, pas l'inverse.

> Problème de famille récurrent, à décider une fois pour toutes : les **constantes de plafond** (`MAX_NODES_PER_SHEET`, `SMALL_RANGE_CELLS`, `MAX_VALUE_SAMPLE`, `MAX_CELLS_PER_SHEET`, `MAX_DENSE_CELLS`) sont lues par **plusieurs** modules qui montent (structure, sweep, graph, resolver). Si chacune est posée dans le module qui la *lit* et ré-exportée ailleurs, on crée des arcs croisés. Recommandation : soit un module `limits.py` pur (feuille), soit une règle « chaque constante vit dans le module feuille qui la définit et les consommateurs l'importent » appliquée strictement. Le plan laisse cette ambiguïté non résolue (il le reconnaît d'ailleurs pour `MAX_VALUE_SAMPLE`, étape 4, en « laisser à l'endroit qui la lit, sinon ré-exporter » — trop flou).

---

## 2. Risques spécifiques

### 2.1 Conversion des fermetures → méthodes de `GraphBuilder`
**Pas de piège de capture au sens classique** (late binding / variable mutée après itération de boucle) : les fermetures capturent des **dicts par référence** (`nodes`, `edges`, `input_nodes`, `cell_owner`…) et n'itèrent aucune variable de boucle en late-binding. Les convertir en `self.nodes`, `self.edges`, … est sémantiquement sûr.

**Le vrai danger n'est pas la capture, c'est l'ORDRE**, parce que trois artefacts du graphe dépendent de la séquence exacte des appels :
1. **Les ids d'arêtes** : `add_edge` calcule `"id": f"e{len(edges)}"` (l.1240) — l'id dépend de la longueur courante du dict, donc de l'ordre de construction. Deux graphes de même topologie construits dans un ordre différent auront des ids `e0/e1/…` permutés.
2. **L'ordre des nœuds** : le retour fait `list(nodes.values())` (l.1583), ordre d'insertion.
3. **L'état partagé du resolver/budget/engine** : le resolver est un **singleton instancié une fois** (l.1040) avant le balayage et le graphe. Chaque `resolver.describe`/`value`/`preload_steps` appelle `budget.take()` (l.458) et **mute l'engine** (`engine.set_formula` l.509, et ré-injection `set_value` dans la chaîne de récupération) et la `_step_cache`. Donc si une phase est appelée deux fois, réordonnée, ou si on ajoute une lecture « en trop », le **budget se déplace** et la **frontière de coupe de décomposition** se déplace → des nodes perdent leur `steps` là où avant elles en avaient.

**Conséquence concrète :** l'ordre de phase doit rester strictement `sélection → defined-names → formules → VBA → Power Query` (le plan le respecte, c'est bien). Mais tout micro-refactor qui touche au *nombre* ou à l'*ordre* des appels resolver (même sans changer la topologie) déplacera la coupe budget et fera échouer un snapshot « octet pour octet » — correct comme tripwire, mais il faut que l'implémenteur sache que c'est une cause fréquente de faux positif et vérifie au lieu de « réordonner pour faire passer le diff ». Voir 2.4 pour le corpus de fixtures.

### 2.2 Le snapshot `scripts/freeze_graph.py` — garantie bonne, mais conditions précises
Le snapshot est **la meilleure idée du plan** et le bon filet pour un refactor qui accepte de toucher aux tests. Mais « JSON canonique complet, identique octet pour octet » n'est vrai que si le script normalise **obligatoirement** :
- `meta.analyzedAt` (l.1559 : `datetime.now(UTC).isoformat()`) et le champ racine `analysisId` (l.1597 : `uuid.uuid4().hex[:16]`) — sinon **jamais** deux runs identiques. Le plan dit « normalisés » sans nommer ces deux champs ; il faut les retirer explicitement.
- Décider du **tri** : trier `nodes`/`edges` par `id` stable rend le diff robuste aux réordonnances d'insertion mais **masque** les régressions qui changent les ids `e0/e1` (2.1). À l'inverse, garder l'ordre d'insertion JSON rend le diff très strict et détectera tout réordonnancement (souhaitable mais source de bruit). **Recommandation :** deux vues — (1) diff « brut ordre d'insertion » qui est la jauge stricte d'équivalence octet, et (2) une vue « topologique triée (id/kind/source/target/kind) » pour juger si un écart est un vrai changement de topologie ou un simple déplacement.

**Champs/variables cachées à contrôler dans le corpus :**
- Ne PAS prendre comme fixtures les classeurs qui **frôlent le budget temps** (ex. `running_totals` de `test_robustness`, `step_seconds=0`) : là, la coupe de décomposition dépend du compte d'appels exact et devient une variable cachée ultra-sensible (2.1). Choisir des classeurs sans plafond atteint.
- Le branche quarantaine (`_quarantine_unresolvable`) produit un warning contenant le texte d'exception `exc` (l.1022) : garantir que le message est déterministe sur la fixture, sinon échec fantôme.
- Le snapshot ne couvre que les fixtures qu'on lui donne : s'assurer d'au moins **un** cas par branche coûteuse — quarantaine (engine), VBA/PQ mocks **ne passent PAS par `analyze_workbook` de façon unitaire** ici, defined-names multi-targets, range grande (branche approx), et un fichier multi-feuilles pour l'ordre `sheet_dims`. Le plan ne dit pas quel corpus précis ; c'est le point le plus à verrouiller.

### 2.3 Dissolution du façade (étape 1) : symboles oubliés / sites runtime
Le plan (§2.5, table des symboles) liste bien les **imports de tête** de fichiers. Mais il manque la moitié des usages réels, qui sont des **sites runtime** et non des imports :

| Site | Ligne | Ce qui casse quand | Conséquence |
|---|---|---|---|
| `monkeypatch.setattr(analyzer, "MAX_CELLS_PER_SHEET", 30)` | test_robustness.py:191 | étape 1 retire l'import `MAX_CELLS_PER_SHEET` d'analyzer | `monkeypatch.setattr` lève `AttributeError` (attribute absent). NB : le patch est **déjà inerte aujourd'hui** (loader lit sa propre globale, cf. loader.py:150/155/208/267) → la ligne peut être supprimée, pas retargetée. |
| `from linexcel.analyzer import _Budget` (dans le corps de tests) | test_robustness.py:269,279,287,296 | étape 4 (resolver) | Runtime import → `ImportError`. Non couvert par le « reformuler les imports de tête ». |
| `monkeypatch.setattr(analyzer, "MAX_CELLS_PER_SHEET"/"SCAN_CHUNK_ROWS", …)` | test_lineage.py:2574,2583 | étapes 3+5 (constantes dans engine, et sweep lit via loader/engine) | Patch inerte → le test de clip (attendu `totalFormulas == 20`) échoue par assertion. Nécessite de retargeter sur `loader.MAX_CELLS_PER_SHEET` et `engine.SCAN_CHUNK_ROWS`, **et** que le code de sweep lise ces constantes via l'attribut module (`loader.MAX_CELLS_PER_SHEET`, pas un `from … import` figé), sinon le patch ne porte pas. |
| `from linexcel.analyzer import SCAN_CHUNK_CELLS, SCAN_CHUNK_ROWS, _chunk_rows` (inline) | test_lineage.py:2592 | étape 3 (engine) | Runtime import. |
| `from linexcel.analyzer import _is_unresolvable` (inline) | test_lineage.py:1155,1171 | étape 3 (engine) | Runtime import. |
| `monkeypatch.setattr(analyzer, "sheet_bytes", …)` | test_cli.py:149,161,171 | étape 2 (cli passe à `from linexcel.structure import sheet_bytes`) | Aujourd'hui ça marche car `cli` fait l'import **dans le corps de la fonction** (cli.py:173), donc l'attribut `analyzer.sheet_bytes` est relu à chaque appel. Dès que cli importe depuis `structure`, il faut retargeter sur `structure.sheet_bytes` (ou `cli.sheet_bytes`). |
| `monkeypatch.setattr("linexcel.analyzer.extract_vba_modules", …)` | test_lineage.py:1563 | étape 6 (VBA → graph) | **Mock par chemin de module.** `analyze_workbook` appelle `extract_vba_modules(...)` via les globals d'analyzer (l.1412). Déplacé dans `graph`, il faut retargeter le path sur `linexcel.graph.extract_vba_modules` (et que graph l'importe comme global). C'est un changement de **corps de test**, pas un import. |
| `monkeypatch.setattr(analyzer, "read_queries", …)` | test_powerquery.py:184,194 | étape 6 (PQ → graph) | Idem : `read_queries` appelé via les globals d'analyzer (l.1480) ; retargeter sur `linexcel.graph.read_queries`. |

**Conclusion étape 1 :** « pur déplacement d'import » est sous-estimé. Il faut lister ces ~11 sites comme des changements de **corps de test** et faire passer la suite complète + snapshot à l'étape 1, pas seulement `pytest` sur 4 fichiers « imports ». Une étape qui laisse des `monkeypatch.setattr` sur des attributs disparus lèvera en premier, donc ce n'est pas *silencieux* — mais c'est plus que « 0 risque ».

### 2.4 Symboles documentés / dépendances de ré-export non couvertes
- `CHANGELOG.md` documente `linexcel.analyzer.sheet_bytes()`. Le plan dissout `sheet_bytes` vers `structure` **sans le ré-exporter** d'`analyzer` (il ne ré-exporte que `inspect_workbook`). Si on tient à la rétro-compat doc/CLI, soit ré-exporter `sheet_bytes` depuis analyzer, soit assumer la cassure explicitement (le plan accepte de casser le privé — à trancher, pas laissé implicite).
- `test_progress.py:16` et `test_cli` importent `inspect_workbook` **depuis analyzer en tête** : ils dépendent du **ré-export** qu'`analyzer` doit conserver (étape 2/7). Cohérent avec la fiche analyzer (« inspect_workbook ré-exporté ») — mais alors il ne faut **pas** l'enlever d'analyzer à l'étape 8 au motif que « les tests doivent pointer vers structure » : le plan est ambigu (table §2.5 dit « maison = analyzer(structure)/structure »). Trancher : soit analyzer ré-exporte inspect_workbook en permanence (et test_progress/test_cli restent valides), soit on déplace aussi ces imports — pas les deux en même temps.

---

## 3. Pièges que le plan aurait ratés (relecture code réel)

1. **Dépendance de module non notée : `tables.py` importe `linexcel.insights`** (runtime). Le graphe d'import cible (§2.5/§4) omet `insights` ; ce n'est pas bloquant (tables est déjà extrait et stable), mais toute carte de dépendances « acyclique » doit inclure `tables → insights` pour être honnête.
2. **`external.py` importe `loader`** (`MAX_DENSE_CELLS`, `declared_cells`). Puisque `structure` dépend d'`external` et de `loader`, c'est cohérent — mais ça signifie que `loader` est un *nœud bas* qu'on ne doit jamais faire dépendre d'`external`/`structure`. Le plan le comprend (note §2.3 structure : « loader ne doit pas dépendre d'external ») — correct.
3. **Le `TYPE_CHECKING` de `decompose` (l.25-26) est déjà le coupe-cycle.** Après l'étape 4 il doit pointer `resolver._ValueResolver` au lieu d'`analyzer._ValueResolver`. Bien vu par le plan (fiche resolver). À condition que `resolver.py` importe `decompose` **en runtime** (`_guard_fallback_expr` l.590, `_scratch_eval` l.465, `SCRATCH_SENTINEL`/`SCRATCH_SHEET` l.500-528, `_render_expr`) et que `decompose` ne garde `resolver` qu'en `TYPE_CHECKING` — sinon cycle `resolver ⇄ decompose`. Vérifier par un `python -c "import linexcel.analyzer"` à chaque étape (ou un check d'import cyclique en CI).
4. **`a1` (l.853) est une dépendance de nommage omniprésente** (ids de nœuds `c:`, `g:`, addresses, samples). Le plan le déplace dans `refs`. Vérifier qu'aucun module/test n'importe `a1` depuis analyzer (aucun trouvé — bon), et que `graph` (où sont tous les appels à `a1`) importe `refs.a1`. Aussi ne pas confondre avec `_a1_position` (l.816) qui, lui, reste un helper resolver/externe.
5. **Effet de bord silencieux de `preload_steps`** : il **vide et mute** la `_step_cache` (l.479) et pose des formules dans la feuille scratch (l.500-512). Si le refactor « extrait » la décomposition dans `graph` mais appelle `resolver.preload_steps` une fois de plus/moins (ex. sur un groupe volatile reclassé), le cache gelé change. Ne jamais faire de la décomposition une fonction « pure » sans transporter le resolver+budget tel quel.
6. **L'ordre `nodes` vs `_stepped` du reporting** : la boucle des nœuds formules est sous `_stepped(...)` (l.1317). C'est un détail de reporter, mais `_stepped` vient de `loader` (importé l.54) ; graph.py devra l'importer depuis loader (le plan l'évoque). Pas bloquant.
7. **`read_external_links`, `find_workbooks`, `resolve_books`, `macro_files`** restent appelés **dans `analyze_workbook`** (l.978-983) et **dans l'orchestrateur final**. `macro_files` est aussi appelé dans le bloc VBA (l.1418) → quand VBA part dans graph, graph doit importer `macro_files` depuis external. Cohérent, mais c'est une 2e occurrence (structure lit les liens pour l'analyse, graph lit les add-ins VBA). Le plan ne distingue pas les deux usages d'`external` entre structure et graph.

---

## 4. Mesures d'atténuation (par risque)

| # | Risque | Mitigation concrète | Vérification |
|---|---|---|---|
| R1 | Étape 1 = faux « 0 risque » : ~11 sites runtime (`monkeypatch`, mock-path, imports inline) | Faire une passe dédiée sur les **sites runtime** (table §2.3), pas seulement les imports de tête. Retargeter `monkeypatch`/`mock.patch` vers le module d'origine réel (loader/engine/structure/graph) ; supprimer les patches déjà inertes (ex. test_robustness:191). | `uv run pytest tests/ -q` après l'étape 1 (suite complète, pas seulement 4 fichiers) + snapshot. Un `grep -rn "analyzer" tests/` doit ne laisser que `analyze_workbook`/`inspect_workbook` légitimes. |
| R2 | Réordonnancement (volontaire ou accidentel) déplaçant budget/ids d'arêtes/ordre nodes | Conserver strictement l'ordre `sélection → names → formules → VBA → PQ`. Ne JAMAIS réordonner dans une étape de « déplacement pur ». Bannir tout memoïsation nouvelle du resolver. | Snapshot brut ordre d'insertion identique à chaque micro-commit de l'étape 6 ; diff topologique trié en appoint pour distinguer « réordonnancement » de « vraie régression ». |
| R3 | Snapshot non déterministe (`analyzedAt`, `analysisId`) ou corpus sous budget | Freeze normalise : retirer `meta.analyzedAt` + `analysisId`. Corpus = fixtures **sans** plafond budget/temps atteint, + 1 cas par branche (quarantaine, grand-range approx, defined-names multi-targets, multi-feuilles, VBA, PQ). | Deux exécutions consécutives du freeze sur la même revision → diff vide. |
| R4 | Mock-path VBA/PQ cassés à l'étape 6 | Retargeter `linexcel.analyzer.extract_vba_modules` → `linexcel.graph.extract_vba_modules` et `analyzer.read_queries` → `linexcel.graph.read_queries` ; veiller à ce que graph importe ces noms comme globals et les appelle en nom nu. | `pytest test_lineage.py -k vba test_powerquery.py -q` verts à l'étape 6. |
| R5 | Unification `resolve_rect_edges`/`_resolve_vba_write` change le graphe (branche grand-range approx absente côté VBA) | Unifier **avec** un flag `allow_approx`/garde de la branche `kept_groups` ; ne pas fusionner naïvement les 3 branches. | Snapshot + `test_lineage.py -k vba` + suite. Si le diff bouge, l'unification est fausse. |
| R6 | `structure → graph` (MAX_NODES_PER_SHEET) | Poser `MAX_NODES_PER_SHEET` (et les plafonds partagés) dans `structure`/`loader`/un `limits.py` et le faire **importer** par graph ; interdire toute flèche `structure → graph`. | Audit d'imports final : aucun module bas n'importe `graph`/`analyzer`. |
| R7 | Cycle `resolver ⇄ decompose` | `resolver.py` importe `decompose` en runtime ; `decompose` garde `resolver` en `TYPE_CHECKING` seulement. | `python -c "import linexcel.analyzer, linexcel.graph, linexcel.resolver"` vert à chaque étape (à mettre en CI/script). |
| R8 | Rétro-compat `linexcel.analyzer.sheet_bytes` (changelog) / imports tests de `inspect_workbook` | Trancher explicitement : ré-exporter `sheet_bytes` et `inspect_workbook` depuis analyzer (recommandé, coût ~2 lignes) OU déplacer les imports de test_progress/test_cli d'un coup à l'étape 8 — mais pas laisser la table §2.5 ambiguë. | Recherche `from linexcel.analyzer import` restante = uniquement les symboles publics voulus. |

---

## 5. Verdict

**Le plan est sain, structuré, réversible, et sa vision d'architecture est correcte et fidèle au code.** Il est **presque prêt**, mais pas tout à fait : il sous-estime systématiquement le travail réel de **retargetage des sites runtime de test** (mocks/patches/imports inline), et laisse deux ambiguïtés d'architecture qui peuvent casser son propre graphe d'import cible (le siting de `MAX_NODES_PER_SHEET`, et la vraie non-identité des deux fonctions qu'il veut unifier à l'étape 7). Ce sont des corrections de plan, pas des changements de comportement — aucune ne remet en cause le découpage ni l'ordre global des étapes.

### Les 3 choses les plus importantes à surveiller pendant l'implémentation
1. **Ordre = sémantique.** Parce que les ids d'arêtes (`e{len(edges)}`), l'ordre des nœuds et surtout l'état du `resolver` (budget + engine muté) dépendent de la séquence exacte des appels, tout déplacement doit être strictement « fermeture → méthode » **sans réordonner**, validé par le snapshot brut à chaque micro-commit. Un écart de snapshot n'est une vraie régression que s'il persiste hors réordonnancement — apprendre à le distinguer avec la vue topologique triée.
2. **Le corpus + la normalisation du snapshot.** Sans retirer `analyzedAt`/`analysisId` et sans fixtures qui n'atteignent pas les plafonds budget/temps, le « octet pour octet » est soit impossible, soit trop sensible — le filet devient du bruit. C'est la condition de confiance de toutes les étapes 2 à 7.
3. **Les ~11 sites runtime de tests (mocks/patches/imports inline)**, qui ne sont ni des imports de tête ni des assertions : ils doivent être retargetés vers le module d'origine réel (loader/engine/structure/graph) au fil des étapes, pas à l'étape 8. C'est la source la plus probable d'échecs « inexplicables » en cours de route (patches devenus inertes ou `AttributeError` sur des attributs disparus).

*Audit rédigé sans modifier aucun code ni committer.*
