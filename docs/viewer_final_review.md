# Contrôle final de convergence — Viewer linexcel

Branche `polish/viewer-impeccable` · après cycle 2 du builder (commit `48b8387`).
Contrôleur final indépendant. **Aucun code source modifié.** Viewer réel généré
via le **package local** `uv run linexcel analyze tests/fixtures/power_query.xlsx
-o /tmp/final.html` puis piloté sous Chromium headless (Playwright) : lecture
DOM + styles calculés, clics réels, clavier (Échap), survol souris réel.

## Verdict final : PRÊT À MERGER ✅

Les trois corrections du cycle 2 sont bonnes et vérifiées en réel ; aucune
régression sur le périmètre demandé. La CI doit rester verte au merge.

---

## 1) B1 — État « Cytoscape manquant » : corrigé, vérifié en réel ✅

Variante `/tmp/nocy.html` = même rendu avec les 5 bundles lib (Cytoscape/dagre/
fcose) retirés du `<head>` → `typeof cytoscape === 'undefined'` à l'exécution.

| Vérif réel | Résultat |
|---|---|
| `#lin-panel` masqué | `class=hidden`, `getComputedStyle().display == "none"` — plus de bande blanche 440 px ✅ |
| `#lin-stats` rempli | `"0 formulas · 5 nodes · 3 edges · 2 Power Query"` (bloc stats posé **avant** le test Cytoscape, l.842-853) ✅ |
| Chrome graph éteint | `#lin-tools` hidden, `#lin-search` disabled, fallback affiché ✅ |
| Erreurs console | aucune (ni pageerror) ✅ |

Cohérent avec l'état « 0 nœud ». Code l.856-873 (`panel.classList.add('hidden')`).

## 2) m1 — Échap dans une recherche vide : corrigé, vérifié en réel ✅

Test runtime fidèle : clic réel sur le nœud `DonnéesExternes_1` (1 nœud sélectionné,
panneau détail ouvert, caméra relevée), focus boîte de recherche **vide**, Échap.

| Vérif réel | Résultat |
|---|---|
| Sélection conservée | `nodes(':selected').length` reste 1 ✅ |
| Panneau non refermé | détail toujours affiché ✅ |
| Caméra intacte | `zoom` ET `pan` **strictement identiques** avant/après ✅ |
| Blur de la boîte | activeElement quitte la recherche ✅ |

Contre-épreuve (le bon chemin est conservé) : recherche **active** puis Échap →
`restoreGraph()` vide bien la sélection (Run: 1 → après Échap: 0). ✅

## 3) m2 — Survol de nœud : corrigé, rendu non cassé ✅

Survol souris réel d'un nœud : la classe `hovered` est ajoutée
(`classes == "name,hovered"`) → bordure renforcée + label réaffiché même en LOD
(`node.hovered` déclaré après `node.far` dans `buildStyle`). Déplacer le curseur
sur le fond du canvas retire `hovered`. **Aucune erreur JS ; le rendu du graphe
et la sélection par clic restent intacts** (les handlers `tap` cohabitent).

Note mineure (non bloquante, cosmétique) : si le pointeur quitte le **canvas
entièrement** (vers le header), `hovered` persiste jusqu'au retour du curseur
sur le fond du graphe, où il se nettoie (vérifié). Autonettoyant ; le code le
commente comme « mouseout à la sortie du canvas » mais ce cas extrême ne déclenche
pas le `mouseout` cytoscape. Sans impact utilisateur réel.

## 4) Régression — verte ✅

- `uv run pytest tests/test_viewer_ui.py tests/test_viewer_values.py -q`
  → **148 passed** (dont les nouveaux tests verrouillant B1/m1, cf. commit 48b8387).
- Viewer normal (`power_query.xlsx`, 5 nœuds) : `#lin-stats` rempli, panneau détail
  actif, 3 canvases, **aucune erreur console/pageerror**.

## Couverture des points de la review (docs/viewer_review_notes.md)

- **B1** → traité et vérifié (voir §1). ✔
- **m1** → traité et vérifié (voir §2). ✔
- **m2** → traité et vérifié (voir §3). ✔
- m3/m4/m5 : hors périmètre des 3 commits (cosmétique / doc limite / wording),
  toujours ouverts mais non bloquants.

*Contrôleur final. Aucun code modifié. Viewer générés : `/tmp/final.html`
(normal), `/tmp/nocy.html` (Cytoscape manquant).*
