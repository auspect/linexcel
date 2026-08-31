# Spec — adversaires & refactorisation de linexcel

## Objectif
Tester le moteur de lineage de linexcel avec des fichiers Excel/VBA **réalistes et hostiles**,
identifier ce qui le fait planter, et **refactoriser proprement** `analyzer.py` sans régression.
Le tout converge vers un **conseil de réflexion** documenté avec preuves visuelles (cellule surlignée).

## Contexte technique (vérifié)
- Engine `formualizer>=0.8.4`. `evaluate_all()` est all-or-nothing : une ref manquante (même IFERROR)
  empoisonne le graphe → `get_value` renvoie None partout. `set_value` write-back + résolution
  dépendance-d'abord (`_resolve_precedents`, `MAX_CHAIN_DEPTH=24`) récupèrent.
- Pas de clock (`TODAY()` → 25569 = 1970). Dates 1900/1904 systèmes connus. Pas d'API number-format.
- `analyze_workbook(data)` (ligne ~859) est une grosse fonction monolithique (~680 lignes) + `_ValueResolver` (~500 lignes). C'est LE code "brut" à refactoriser.
- Pre-commit gate : `uv run ruff check` + `uv run ruff format` + `uv run pytest` (590 passed).
- Budgets existants : `MAX_*`, `MAX_SCRATCH_EVALS`, `_Budget`.

## Livrables (3 axes disjoints, 3 agents)

### Axe A — kimi : générateur de fixtures + probe de crashs
Créer dans `tools/` :
- `gen_fixtures.py` : génère des .xlsx/.xlsm **réalistes** avec taille **paramétrable**
  (nb lignes/colonnes/feuilles via args CLI). Couvre des cas hostiles :
  - fichiers avec macros VBA (.xlsm) — macros présentes mais inoffensives, et une macro
    "hostile" qui échoue (erreur VBA) pour tester le path VBA
  - **macro security blocks** : fichier signé/verrouillé, conteneur Office Open XML avec
    `vbaProject.bin` chiffré / absent
  - **étiquettes de confidentialité** : propriétés document (company/classification dans
    `docProps/core.xml` + `app.xml`), le genre que xlwings manipule
  - dates système 1904, refs cassées `=NOSHEET!A1`, formules croisées multi-feuilles,
    `IFERROR(NOSHEET)`, `=SUM(A1:A7)` sur formules, self-referencing cycles
- `probe_crashes.py` : exécute `analyze()` sur chaque fixture générée, capture
  succès/exception/None-values, et écrit `crash_report.md` (table : fixture, phase, résultat,
  symptôme). C'est LA source de vérité des crashs.
- Fixtures générées déposées dans `tests/fixtures/adversarial/`.

### Axe B — claude : refactoriser `analyzer.py` proprement
Découper la monolithique `analyze_workbook` et `_ValueResolver` en fonctions/classes cohérentes,
lisible, sans changement de comportement. **Aucune régression** : les 590 tests doivent rester verts.
Ajouter si pertinent des tests unitaires qui verrouillent la nouvelle structure.

### Axe C — agy : conseil de réflexion avec preuves visuelles
Rédiger `docs/adversarial-reflection.md` : analyse des cas hostiles, lesquels plantent et pourquoi,
comment les gérer intelligemment (stratégies de résilience), et le plan de refactorisation.
Inclure des **preuves visuelles** : captures de la cellule surlignée (viewer HTML) via pixelrag
(`pixelshot` / `pixelrag`). Documenter les fichiers/screenshots produits.

## Règles pour chaque agent
- **Écris ton analyse DANS le fichier livrable dès le début** (mandat d'écriture précoce), puis enrichis.
  Si tu dois t'arrêter, le fichier doit déjà exister et être commité.
- Reste sur TON axe — ne touche pas aux fichiers des autres axes.
- Commit à chaque étape. Préfixe de commit : `axe-<lettre>: ...`.
- Réponds en français, concis.

## Critères d'acceptation globaux
- `uv run pytest` vert sur la branche mergée (590 + nouveaux, pas de régression).
- `uv run ruff check` + `ruff format` propres.
- `crash_report.md` réel (pas d'invention), `docs/adversarial-reflection.md` présent avec images.
- Branches mergées dans `adversarial-refactor`.
