# Mission : performance linexcel — engine init + evaluate_all

Repo : /home/phili/linexcel, branche fix/viewer-and-perf (un fix viewer y est déjà commité, ne le casse pas).

## Problème signalé par l'utilisateur
Sur un fichier réel, l'analyse a pris **>3000 secondes** dont **2900s pour "engine init" + "evaluate all"**, alors que l'estimation affichée était "~3 minutes". L'estimation (`src/linexcel/cli.py`, `SECONDS_PER_SHEET_MB`, formule `weight/1MB * SECONDS_PER_SHEET_MB`) est un **plancher basé sur la quantité de formules**, qui ignore la complexité des dépendances (chaînes longues, références croisées, feuilles volumineuses). Donc l'estimation sous-estime gravement.

## Objectif
1. **Comprendre** pourquoi engine init + evaluate_all peut prendre 2900s alors que l'estimation en prévoit 3min. Identifier le goulot (boot_engine dans engine.py, evaluate_all de formualizer, quarantaine _quarantine_unresolvable, etc.).
2. **Vérifier la doc/API de formualizer** (pip: formualizer) pour voir si evaluate_all est mal utilisé (ex: un paramètre qui désactive une optimisation, un mode de recalcul naïf, un timeout manquant, etc.). Chercher s'il y a un meilleur pattern (évaluation incrémentale, par sous-ensemble, lazy).
3. **Améliorer l'estimation** : elle doit refléter au moins l'ordre de grandeur réel. Une heuristique meilleure (nombres de formules × profondeur moyenne de chaîne, ou coût estimé par formule) vaut mieux que le seul poids.
4. **Suggestions en console quand ça dure trop longtemps** : afficher un avertissement actionnable pendant le run quand le temps dépasse l'estimation d'un facteur, avec des pistes (--time-budget, -v, couper les chaînes, etc.).

## Comment travailler
- Profil sur des fichiers dispo : `benchmarks/...` n'existe pas ici, mais `tests/fixtures/` a des xlsx. Tu peux aussi générer un gros fichier avec openpyxl (MAIS openpyxl ne met pas de cache de formule — note-le). `uv run python -m cProfile -s cumulative ...` sur une analyse.
- Lis la doc formualizer : `uv run python -c "import formualizer, inspect; print(formualizer.__file__)"` puis cherche evaluate_all / la classe engine. Vérifie sur PyPI/GitHub formualizer la doc du moteur.
- Mesure avant/après chaque changement.

## Règles
- Écris et commite tôt (git add + commit sur fix/viewer-and-perf). Ne push pas.
- `uv run pytest tests/ -q` doit rester vert après chaque changement.
- Ne change PAS l'API publique (analyze_workbook, LineageResult).
- Si le fix de perf est trop risqué/gros pour ce run, livre au minimum : l'analyse de cause racine documentée + l'estimation améliorée + les suggestions console (les 3 changements sûrs), et documente le fix perf risqué comme prochaine étape.

## Livrable
- Rapport français : cause racine du 2900s, ce qui a été changé (estimation, suggestions, éventuel fix perf), mesures avant/après, tests verts. Si tu as identifié une mauvaise utilisation de formualizer, documente-la précisément.
