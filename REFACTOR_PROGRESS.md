# État du refactor linexcel (pour reprise)

Dernière mise à jour: 2026-09-02

## Où en est-on
Refactor de linexcel (github.com/auspect/linexcel) sur la branche `refactor/analyzer-orchestrator` du clone `/tmp/linexcel_refactor`.

Objectif : `src/linexcel/analyzer.py` (1882 lignes monolithe) → orchestrateur mince (~110-150 lignes).

## Progression
- **Étapes 1-6 COMMITÉES et vertes (586 tests)** :
  - step 1 : façade de ré-export dissous (tests)
  - step 2 : `structure.py` extrait
  - step 3 : `engine.py` extrait + câblé (boot_engine, EngineSession, quarantaine)
  - step 4 : `resolver.py` extrait (_Budget, _ValueResolver)
  - step 5 : `sweep.py` extrait (FormulaGroup, extraction+regroupement R1C1)
  - step 6 COMPLÈTE (6/6 sous-commits) : `graph.py` GraphBuilder — sélection nœuds, edge/node infra, build_names, build_formula_nodes, build_vba, build_queries
- **Étape 7 EN COURS (non commitée)** : réécriture de analyze_workbook en orchestrateur mince. analyzer.py déjà à **208 lignes** (de 1882). Fichiers modifiés non commités : analyzer.py, graph.py, powerquery.py.
- **Étape 8** : nettoyage imports privés restants + vérif qu'aucun module métier n'importe analyzer.

## Reprendre
Le brief prêt pour relancer est `/tmp/refactor_brief.md` (étapes 7-8).
Commande : `claude -p "Lis /tmp/refactor_brief.md puis exécute exactement ce qu'il demande" --allowedTools "Read,Write,Edit,Bash" --max-turns 90`
Méthode : sprints de 90 turns (Claude sature à chaque fois), vérifier `uv run pytest tests/ -q` = 586 passed + snapshot graphe identique à chaque étape.

## Références
- Plan complet : `/tmp/linexcel_refactor/REFACTOR_PLAN.md`
- Revue croisée Kimi : `/tmp/linexcel_refactor/KIMI_REVIEW.md`
- ~11 sites runtime de tests à retargeter au fil des étapes (monkeypatch.setattr(analyzer,...) test_robustness:191, test_lineage:2574/2583, test_cli:149/161/171; mock-paths linexcel.analyzer.extract_vba_modules test_lineage:1563, analyzer.read_queries test_powerquery:184/194; imports inline test_lineage:1155/1171/2592, test_robustness:269)

## Note
Dernier run Claude Code stoppé par "session limit resets 7pm UTC" — pas un bug de code.
