# Revue qualité — Viewer linexcel (branche `polish/viewer-impeccable`)

Contrôleur qualité indépendant. Évaluation de l'état ACTUEL de
`src/linexcel/assets/viewer.html` après les trois commits du builder
(`2461819`, `9693c4a`, `00c6d5a`). Aucun code source n'a été modifié.

## Verdict d'ensemble

**Bon, solide, proche du "impeccable". Aucun point bloquant.**
Les trois changements du builder sont justes dans leur intention et bien
exécutés ; les états vides et d'erreur sont proprement éteints (chrome graph
mis à l'écart, messages clairs). Deux imperfections réelles subsistent, dont
une dans le nouvel état d'erreur Cytoscape lui-même. Le reste est finition
mineure et suggestions. Le viewer est publiable tel quel.

### Ce qui a été vérifié en réel
- **Régression** : `uv run pytest tests/test_viewer_ui.py tests/test_viewer_values.py -q`
  → **146 passed**. Aucune erreur JS console sur aucun cas (seul un warning
  bénin Cytoscape `wheel sensitivity`, à ignorer).
- Rendus réels générés via le package **local** (`uv run linexcel analyze` — voir
  piège ci-dessous) : classeur normal `power_query.xlsx` (5 nœuds), graphe dense
  généré (168 nœuds / 232 arêtes), classeur sans formule `static.xlsx` (0 nœud),
  et variante Cytoscape manquant (bundles retirés du `<head>`).
- Le tout piloté sous Chromium headless (Playwright) : clics, clavier (Échap,
  onglets), changement de thème, toggle de layout, filtres de type, recherche,
  lecture DOM + styles calculés. Captures d'écran sauvegardées pour contrôle
  visuel : `/tmp/v_norm_local.png`, `/tmp/v_big_local.png`, `/tmp/v_empty_local.png`.
  *(La vision par image n'était pas accessible dans cet environnement : le jugement
  visuel repose sur styles calculés/lecture CSS, pas sur une relecture pixel.)*

## Axes d'évaluation

### 1) États vides / erreurs — bien, une imperfection
- **Graphe 0 nœud (classeur sans formule)** — propre. Message centré
  « No formulas were found in this workbook, so there is no lineage graph to
  display. » (i18n dans les 10 langues), chrome graph éteint : `.lin-tools` et
  `.lin-legend` `hidden`, recherche et filtre de feuille `disabled`, groupes rail
  « Node types » et « Layout » retirés, panneau détail masqué. L'onglet **Sheets
  reste vivant et complet** (liste + préview des cellules rendues) : le rapport
  reste navigable. Bon.
- **Erreur Cytoscape manquant** — correct mais **incomplet** : le fallback
  s'affiche et le chrome graph est éteint (comme la branche 0 nœud), **mais le
  panneau détail `#lin-panel` (440 px) reste visible et vide**, et la ligne stats
  du bandeau reste blanche. Voir point **B1**.
- Chaque onglet/paneau : les panneaux Overview/Screenshots/Sheets s'affichent
  seulement si la donnée est là ; aucun blanc parasite observé.

### 2) Finition chrome — très bonne
Header/identité/stat, rail de navigation à gauche, barre d'outils flottante sur
le canvas, panneau détail à droite : logique de 3 zones cohérente, espacée,
typographie homogène (échelle 0.72–0.95rem). États hover présents sur onglets,
boutons d'outils, bouton thème, bouton ✕ ; focus visible global
(`:focus-visible` outline `--focus`, anneau sur la recherche). Rien de criant.

### 3) Graphe dense
Testé sur 168 nœuds/232 arêtes : au-delà de 120 nœuds le layout bascule bien sur
« Flow » (dagre) par défaut, bascule Organic↔Flow fluide (aria-pressed à jour),
filtres de type fonctionnent, légende ne liste que les types présents (2
cavets). Le graph rend sans erreur. LOD (labels estompés en dézoom) raisonnable.
Point d'amélioration mineur : aucun indice au survol d'un nœud (voir **m2**).

### 4) Accessibilité — forte
- **Contrastes** (calculs WCAG) — textes clairs passent tous ≥4.5:1 en clair
  (ink 19.2, ink2 7.7, muted 5.0, link 5.9, focus 5.9, sémantiques ok/danger ≥6.1)
  et en sombre (≥7.3 partout). Les couleurs de palette <4.5 (green 2.7, amber
  2.1…) sont des **remplissages de nœuds**, pas du texte — le texte de badge est
  dérivé par `onColor()` (testé), l'encre des liens ≠ remplissage nœud. Conforme.
- aria correct : pattern `tablist`/`tab`/`tabpanel`, `aria-selected`,
  tabindex roving, `aria-pressed` sur layout/type/thème, `role=region` +
  `aria-label` sur le panneau, `aria-live` sur le statut de recherche, noms de
  chaque contrôle icône-only. Clavier : flèches/Home/End parcourent les onglets
  (seulement les visibles), Entrée/Space natifs.
- **Échap** (changement du builder) : vérifié en réel —
  (a) Échap avec focus sur le corps désélectionne un nœud et referme le panneau
  ✓ ; (b) Échap dans la recherche ne fait pas double déclenchement (la recherche
  possède Échap via son propre handler, l'exemption du handler document joue) ;
  (c) Échap sans sélection : ne fait rien, pas d'erreur ✓.
- Point faible connu et assumé : le canvas Cytoscape n'est pas navigable au
  clavier ni lisible par un lecteur d'écran (les nœuds sont de la peinture
  canvas). Les données restent atteignables par la recherche / les onglets
  Overview/Sheets. À documenter comme limite (voir **m4**).

### 5) Professionnalisme visuel
Haut niveau : hiérarchie claire (header → rail → contenu), palette CVD-safe
déclarée en tokens, thème sombre en opt-in (explicite et motivé dans le code),
aucune fuite de style inline dépendante des données (testé). Cohérent d'ensemble.

### 6) Régression — verte
146 tests verts, aucune erreur JS (normal, dense, vide, sans-Cytoscape), thème
sombre/clair sans erreur, bascule d'onglets propre.

## Piège d'évaluation (à signaler au builder)
`uvx linexcel analyze …` tire le **paquet publié sur le registre**, pas le
template local édité sur la branche — les viewers générés ainsi ne contenaient
pas les ids des nouveaux états. Il faut utiliser `uv run linexcel analyze …`
(package local) pour valider les changements de la branche. Les tests, eux,
utilisent bien `render_html` local. Si le builder valide manuellement avec
`uvx`, il risque de tester une version obsolète.

---

# Points à corriger — classés par priorité

## Bloquant
*Aucun.*

## Important

### B1 — État d'erreur « Cytoscape manquant » : panneau vide et stats absentes
- **Quoi** : dans la branche `if (typeof cytoscape === 'undefined')`
  (viewer.html ~l.839-853), le fallback s'affiche et le chrome graph est éteint,
  mais `#lin-panel` (colonne 440 px) reste visible et **vide** (classe
  `lin-panel lin-empty`, aucun contenu, car `renderPlaceholder()` n'est jamais
  appelé — le `return` a lieu avant, l.852), et `#lin-stats` reste vide.
- **Où** : viewer.html, branche Cytoscape-manquant (l.839-853).
- **Pourquoi** : le message d'erreur du canvas est encadré d'une large colonne
  blanche qui se lit comme un rendu cassé, et le bandeau perd sa ligne de stats.
  La branche « 0 nœud » (l.896-904) gère cela correctement
  (`panel.classList.add('hidden')` après l'avoir posé plus haut les stats) ;
  les deux branches « rien à afficher » devraient être cohérentes.
- **Correction suggérée** : dans la branche Cytoscape-manquant, ajouter
  `document.getElementById('lin-panel').classList.add('hidden')` (ou le masquer
  équivalent), et poser la ligne stats (déplacer le bloc stats l.861-872 avant le
  test Cytoscape, ou appeler le même code) pour qu'elle s'affiche comme en état
  0-nœud.

## Mineur

### m1 — Échap dans une recherche vide désélectionne une sélection non liée
- **Quoi** : quand un nœud a été sélectionné par clic sur le canvas puis que
  l'utilisateur tabule vers la recherche et presse Échap (même recherche vide),
  la sélection est effacée **et** la vue se refit.
- **Où** : viewer.html, handler Échap de la recherche l.1043-1046 →
  `restoreGraph()` (l.1010-1015) qui appelle toujours `clearSel(cy)` + `fit`.
- **Pourquoi** : Échap dans la recherche devrait réinitialiser la **recherche** ;
  effacer une sélection provenant d'un clic (et refitter la caméra) est un
  effet de bord surprenant quand la boîte est vide. Comportement défendable pour
  annuler une recherche active, excessif sinon.
- **Correction suggérée** : dans le handler de la recherche, ne faire
  `restoreGraph()` que si `search.value` est non vide (sinon simplement effacer /
  blurer), ou distinguer « annuler la recherche » (reset query + désélection des
  matchs) de « quitter la boîte » (ne pas toucher au canvas).

### m2 — Graphe dense : aucun indice de survol d'un nœud
- **Quoi** : pas de `title`/emphase au survol ; à faible zoom les labels sont
  estompés (LOD, l.926-934), donc un petit nœud dense est difficile à identifier
  sans le sélectionner.
- **Où** : style/handlers de rendu nœuds (Cytoscape).
- **Pourquoi** : en vue dense, le survol est le geste le plus naturel pour
  « est-ce bien cette cellule ? » avant de cliquer. Simple emphase au survol
  (agrandir légèrement/épaissir la bordure + éventuellement afficher le label
  complet) suffirait. Optionnel ; le panneau de détail et la recherche restent
  les chemins principaux.

### m3 — Rail « orphelin » quand les groupes sont éteints
- **Quoi** : en états 0-nœud / Cytoscape-manquant, seuls « Views » restent dans
  le rail ; l'espace vertical sous le groupe reste vide sur une hauteur de
  rail-plein.
- **Où** : rail latéral, groupe Views (markup l.578-591) une fois Node
  types/Layout masqués.
- **Pourquoi** : cosmétique. À confirmer visuellement que ça ne fait pas
  « vide intentionnellement » sur petit contenu ; acceptable tel quel.

### m4 — Le canvas graph n'a pas d'équivalent accessible dédié
- **Quoi** : le graphe Cytoscape est de la peinture canvas : illisible par un
  lecteur d'écran, non navigable au clavier nœud à nœud.
- **Où** : rendu graph ; les messages fallback/0-nœud sont des `<div>` texte
  bruts (pas de rôle/annonce).
- **Pourquoi** : à documenter comme limite assumée (les données sont atteignables
  par la recherche, Overview et Sheets) plutôt que comme bug ; éventuellement
  ajouter un `role`/annonce aux messages d'état vides/erreur pour qu'un lecteur
  d'écran les signale. Très faible priorité.

### m5 — Vérifier la visée du message stats « 0 formulas » sur un classeur à requêtes
- **Quoi** : sur `power_query.xlsx`, le bandeau affiche « 0 formulas · 5 nodes ·
  3 edges · 2 Power Query » alors que le graphe montre des arêtes de lignage.
- **Où** : stats (`#lin-stats`), l.861-872.
- **Pourquoi** : « formulas » compte les cellules **à formule Excel**, pas les
  requêtes/tables (ici 0 cellule calculée mais un lignage Power Query) : le
  chiffre est probablement correct mais peut surprendre. À vérifier que le libellé
  reste honnête pour un classeur 100 % Power Query ; sinon ajuster le wording.
  Hors périmètre des 3 commits.

---

*Rédigé par le contrôleur qualité indépendant. Aucun code modifié.*
