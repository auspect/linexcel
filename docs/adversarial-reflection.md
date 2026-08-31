# Conseil de Réflexion & Analyse Adversariale de Linexcel (Axe C)

## 1. Objectifs & Cadre de l'Axe C

L'objectif de l'Axe C est de formaliser une analyse rigoureuse de la résilience du moteur de lignage `linexcel` face à des fichiers Excel/VBA complexes, hostiles ou atypiques rencontrés en entreprise. Ce document synthétise :
- La cartographie des scénarios adversariaux réels ;
- L'analyse des mécanismes de défaillance (ce qui plante et pourquoi) ;
- Les stratégies de résilience intelligentes et robustes ;
- Le plan de refactorisation de l'architecture (`analyzer.py` et `_ValueResolver`) ;
- Des **preuves visuelles** concrètes générées via `save_html` et capturées avec `pixelshot` / `pixelrag`.

---

## 2. Cartographie des Cas Hostiles

Les classeurs Excel en environnement de production présentent fréquemment des anomalies structurelles ou logiques :

| Famille de cas hostiles | Exemples types | Risque / Impact potentiel |
|---|---|---|
| **Références cassées & Inexistantes** | `=NOSHEET!A1`, `=#REF!`, `=#NAME?`, `=IFERROR(NOSHEET!A1, 0)` | Empoisonnement du graphe d'évaluation `formualizer` (`evaluate_all()` renvoie `None` globalement). |
| **Cycles & Auto-références** | `=A1+1` dans `A1`, cycles multi-cellules `A1 -> B1 -> A1` | Boucles infinies de résolution ou blocage du résolveur de graphe. |
| **Sécurité VBA & Macros (.xlsm)** | Macros bénignes, macros avec erreurs d'exécution, `vbaProject.bin` absent/chiffré/verrouillé | Échec d'extraction `oletools`, plantages de parsing de flux binaire ou corruption de métadonnées. |
| **Étiquettes de Confidentialité & Propriétés** | Sensibilité / Classification dans `docProps/core.xml` et `app.xml`, custom XML parts | Altération de la lecture OOXML ou rejet par des validateurs stricts. |
| **Systèmes de Dates & Absence d'Horloge** | Date 1904 vs 1900, fonctions volatiles temporelles (`TODAY()`, `NOW()`) | Décalage de 4 ans et 1 jour (1462 jours) ou fallback sur époque Unix (25569 = 1970). |
| **Bombes de Dimensions ("Stray Corners")** | Clic accidentel en `XFD1048576`, `<dimension ref="A1:XFD1048576"/>` | Tentative d'allocation matricielle dense (512 GiB) causant un crash `SIGABRT` OOM dans les liaisons Rust. |
| **Formules Composites & Agrégations** | `=SUM(A1:A7)` sur cellules calculées, formules imbriquées multi-niveaux | Explosion combinatoire des décompositions et épuisement de la mémoire de travail scratch. |

---

## 3. Diagnostic des Pannes : Qu'est-ce qui plante et pourquoi ?

### 3.1. L'évaluation "All-or-Nothing" de Formualizer
Le moteur Rust `formualizer>=0.8.4` utilise un graphe de dépendances strict. Lorsqu'une formule référence une feuille inexistante (`=NOSHEET!A1`), même enveloppée dans `IFERROR()`, l'appel `evaluate_all()` échoue globalement et empoisonne le graphe : chaque cellule interrogée via `get_value()` renvoie `None`.
- **Pourquoi :** Le moteur tente de compiler et résoudre statiquement l'intégralité du graphe avant toute évaluation partielle.
- **Remède actuel :** Le write-back via `set_value()` et la marche dépendance-d'abord (`_resolve_precedents`) permettent de restaurer localement des valeurs calculables.

### 3.2. Le piège de l'allocation dense sur dimensions corrompues
Certains parseurs de haute performance (comme calamine) allouent une matrice dense basée sur la balise `<dimension ref="..."/>`. Une feuille avec une seule cellule en `A1` et une en `XFD1048576` revendique 17 milliards de cellules :
- **Pourquoi :** L'allocation de 512 Go de RAM échoue immédiatement en Rust, provoquant un `panic!` non récupérable en Python (`SIGABRT`).
- **Remède actuel :** Pré-lecture légère du XML (`declared_cells`) et bascule vers `openpyxl` en mode lecture paresseuse (lazy streaming) au-delà de `MAX_DENSE_CELLS`.

### 3.3. Décomposition des formules complexes et sentinelles
Lors de la décomposition pas-à-pas des formules dans la feuille temporaire `__lineage_scratch__`, si le moteur échoue à évaluer une sous-expression, il conserve silencieusement la valeur précédente de la cellule :
- **Pourquoi :** Risque de propager une fausse valeur issue d'une évaluation antérieure.
- **Remède actuel :** Injection de la sentinelle `SCRATCH_SENTINEL = "__linexcel_no_value__"` avant chaque évaluation.

---

## 4. Stratégies de Résilience Intelligente

1. **Isolation par Budgets Stricts :**
   - Plafonds de profondeur : `MAX_CHAIN_DEPTH = 24`.
   - Plafonds d'évaluations scratch : `MAX_SCRATCH_EVALS = 4_000`.
   - Plafonds de nœuds par feuille : `MAX_NODES_PER_SHEET = 400`.
2. **Résolution incrémentale Dépendance-d'Abord :**
   - Visite récursive des antécédents immédiats avant de tenter l'évaluation d'un nœud dépendant.
3. **Tolérance aux erreurs et Fallbacks gracieux :**
   - Récupération des valeurs en cache (valeurs pré-calculées sauvées par Excel dans le fichier XML) lorsque le moteur Rust ne peut pas évaluer.
   - Marquage explicite des avertissements (`result.warnings`) sans interrompre le pipeline utilisateur.

---

## 5. Plan de Refactorisation Architecturale

Le fichier monolithique `analyzer.py` (~2370 lignes) et la classe `_ValueResolver` (~500 lignes) nécessitent un découpage modulaire garantissant 100% de non-régression sur les 590+ tests :

### Architecture Cible Découplée :
- `linexcel.core.structures` : Modèles de données (`CellNode`, `FormulaGroup`, `LineagePayload`).
- `linexcel.core.loader` : Détection des dimensions, inspection des balises XML, streaming openpyxl / calamine.
- `linexcel.engine.formualizer_adapter` : Encapsulation du runtime Rust, gestion de `SCRATCH_SHEET` et sentinelles.
- `linexcel.engine.resolver` : Logique de résolution de dépendances, chaînage de précédents et write-back.
- `linexcel.vba.extractor` : Découplage de l'analyse statique VBA / macros et gestion des erreurs de conteneurs.
- `linexcel.viewer.html_builder` : Génération du viewer HTML autonome Cytoscape.js.

---

## 6. Preuves Visuelles & Démonstration

*(En cours d'enrichissement avec captures réelles via `pixelshot` et `/tmp/agy-demo.xlsx`)*
