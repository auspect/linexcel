# Mandat — Builder viewer linexcel

Tu travailles dans le repo git `/home/phili/linexcel`, branche `polish/viewer-impeccable`. C'est le package PyPI linexcel (analyse de lignage Excel). Le fichier à améliorer : `src/linexcel/assets/viewer.html` (1945 lignes, un template autonome HTML+CSS+JS inline qui génère le dashboard interactif). Il est chargé par `src/linexcel/viewer.py` (render_html) et la sortie est `uvx linexcel analyze f.xlsx`.

## Contexte (vérifié)
- Le viewer a DÉJÀ un design system soigné : tokens CSS (`.lin-root`), thème clair/sombre opt-in, palette de nœuds CVD-safe, navigation par rail latéral (déjà migré), recherche, panneau latéral de détail, accessibilité (rôles, aria, focus-visible). 49 tests UI passent (tests/test_viewer_ui.py, test_viewer_values.py). Un mockup de redesign existe (docs/dashboard_proposal.html) mais le viewer a évolué au-delà et le mockup n'est PAS la spec impérative.
- Objectif : rendre le dashboard **impeccable, propre et professionnel** — peaufiner sans rien casser.

## Comment valider (obligatoire, sinon tu travailles à l'aveugle)
1. `uv run pytest tests/ -q` — doit rester 100% vert (587+ tests, dont 49 viewer). C'est LE filet de sécurité : les tests test_viewer_ui.py et test_viewer_values.py verrouillent le comportement.
2. Captures visuelles : `uv run python scripts/capture_viewer.py` (si playwright installé) génère imgs/*.png. Si playwright n'est pas installé, fais : `uv pip install playwright && playwright install chromium` puis capture. SINON, génère un HTML réel et ouvre-le (tu peux lancer un `python3 -m http.server` en fond et utiliser un outil navigateur/screenshot si tu en as). Au minimum, lis le viewer.html et les tests pour comprendre l'UI.

## Ce que je veux — polir pour "impeccable" (par priorité)
1. **Cohérence et finition du chrome actuel** : header, rail latéral, boutons, focus. Vérifie espacements, alignements, tailles, états hover/focus/actif cohérents. Pas de valeurs magiques incohérentes.
2. **États vides et erreurs** : chaque panneau/onglet doit avoir un état "vide" propre et explicite (pas un blanc vide qui semble un bug). Charge lente / pas de donnée → message clair. Le code commente déjà ce principe ("an empty cell reads as a rendering bug") — étends-le aux panneaux/onglets/graph vide.
3. **Graph** : légende claire, zoom/panneaux, tooltip de nœud lisible, surcharge graphique sur beaucoup de nœuds (le code récent gère déjà les graphes denses — vérifie que ça reste propre).
4. **Accessibilité** : contrastes (les tokens sont déjà choisis pour 4.5:1), focus visible partout, aria sur les contrôles interactifs, navigation clavier (tabs, fermer panneau, recherche).
5. **Professionnalisme visuel** : typographie lisible, hiérarchie claire, espace blanc maîtrisé. PAS de refonte complète — le viewer est déjà bon. Sois chirurgical : améliore ce qui est imparfait, ne réécris pas ce qui marche.

## Regles STRICTES
- **Ne change pas l'API** : render_html, structure du JSON graph, classes CSS utilisées par les tests, strings i18n (dans i18n.py), ids référencés par tests/test_viewer_ui.py. Les tests définissent le contrat.
- **Écris dès maintenant** : si tu identifies une amélioration, applique-la tout de suite dans viewer.html, ne stocke pas tes intentions pour la fin. Commite régulièrement : `git add src/linexcel/assets/viewer.html && git commit -m "viewer: ..."` (git identity déjà configurée).
- Teste APRÈS chaque groupe de changements : `uv run pytest tests/test_viewer_ui.py tests/test_viewer_values.py -q` doit rester vert. Puis la suite complète en fin de run.
- Ne push PAS. Ne change que viewer.html (et éventuellement viewer.py si strictement nécessaire pour un état d'erreur, en gardant render_html compatible).
- Si tu bloques sur quelque chose d'ambigü, documente-le dans un fichier `docs/viewer_polish_notes.md` plutôt que de deviner une direction destructrice.

## Livrable
- viewer.html amélioré, tests verts, commits propres.
- Un court rapport final (français) : ce que tu as amélioré, les choix, ce qui reste ouvert.
