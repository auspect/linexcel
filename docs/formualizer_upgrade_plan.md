# Formualizer : audit et amélioration continue

## Addendum du 7 octobre 2026 : CI sur formualizer 0.11.0

Le workflow de veille a résolu formualizer 0.11.0, publié sous le tag amont
[`v0.11.0`](https://github.com/PSU3D0/formualizer/blob/v0.11.0/CHANGELOG.md).
L'index de recherche ne montrait pas encore cette version. Le changelog tagué,
les stubs Python du même tag et son guide CLI servent de sources :
[stubs Python 0.11.0](https://github.com/PSU3D0/formualizer/blob/v0.11.0/bindings/python/formualizer/formualizer_py.pyi),
[guide CLI 0.11.0](https://github.com/PSU3D0/formualizer/blob/v0.11.0/docs/cli.md).

### État de la CI de PR #113

Le [run 37619156812](https://github.com/auspect/linexcel/actions/runs/37619156812)
a réussi tous les jobs, dont les tests Linux/Windows et navigateur, après le
rebase sur linexcel 1.13.0. Le premier run
[plancher/verrou 37619156894](https://github.com/auspect/linexcel/actions/runs/37619156894)
a passé les 236 tests sur chacun de ces deux canaux. Son job « dernière version »
sur 0.11.0 a révélé
116 échecs et 120 succès, en raison de collisions entre les nouveaux noms
statistiques natifs et les alias linexcel.

Le correctif conserve `NORMDIST` natif après vérification de ses contrats.
`NORMSDIST` garde son bridge : le natif 0.11.0 accepte à tort un argument
supplémentaire. L'enregistrement autorise explicitement cette surcharge ;
les erreurs inattendues restent visibles. Les contrats couvrent les arités,
coercitions, tableaux, erreurs, gardes et recalculs après modification d'entrée.

Les probes ont aussi confirmé les limites du parseur au chargement. Les
formules rejetées pour leur hauteur, imbrication ou budget de taille sont
isolées avant l'import. Les autres cellules restent calculables et les
formules isolées conservent leur cache lorsqu'il existe. Cette admission
dépend du refus réel du parseur : une chaîne de 700 termes continue donc
à être calculée avec 0.10.1. La matrice hebdomadaire couvre désormais les
contrats des alias et les limites natives dans des sous-processus.

La garde vérifie nativement les formules de plus de 64 caractères avant
l'import, en plus du contrôle lexical contre les conversions dangereuses.
Cette protection ajoute du travail de parsing. Son coût sur les grandes
familles de formules reste à mesurer dans le lot de performances ; aucun
gain de vitesse de linexcel n'est revendiqué dans cette release.

Les résultats de 1 453 tests locaux et 99 tests navigateur consignés plus bas
datent de l'arbre avant le rebase sur linexcel 1.13.0. Ils restent un relevé
historique, pas une validation de 0.11.0. Le statut final dépend du nouveau run
après correction des alias et des gardes du parseur.

### Capacités 0.11.0 et priorités d'intégration

| Capacité amont | Situation linexcel | Prochaine action |
| --- | --- | --- |
| Parseur borné : hauteur AST par défaut de 256, source de 65 536 octets, 16 384 tokens, 8 192 nœuds et 72 cadres Pratt | La garde reconnaît les refus de complexité natifs avant import et conserve son plafond contre les conversions profondes. Les stubs Python 0.11.0 n'exposent pas `ParserLimits`. | **P0 livré.** Couvrir les limites dans la veille ; ne pas augmenter les budgets sans API publiée et preuves de sûreté. |
| Fonctions natives Excel 2007, dont `NORMSDIST`, `NORMDIST`, `NORMSINV`, `TDIST` et les fonctions historiques de lois | Le natif `NORMDIST` remplace son bridge ; `NORMSDIST` garde la validation d'arité linexcel. | **P0 livré.** Étendre les oracles statistiques indépendants avant de retirer d'autres adaptations. |
| Opérateur `^` désormais associatif à gauche; `AND`, `OR` et `XOR` évaluent tous leurs arguments et propagent la première erreur selon l'ordre source | Ce sont des changements sémantiques visibles. `2^3^2` vaut maintenant 64; `AND(FALSE,1/0)` renvoie `#DIV/0!`. | **P1.** Ajouter ces cas aux contrats versionnés et comparer les résultats à Excel avant de traiter un écart comme une régression ou une adaptation linexcel. |
| `IF` applique les conditions tableaux élément par élément; tableaux CSE historiques, références de débordement `A1#` et `_xlfn.ANCHORARRAY(A1)` | L'analyse sait lire les formules et les cellules de dépendance, mais ces nouveaux résultats de tableaux et références de débordement n'ont pas encore de contrat linexcel dédié. | **P1.** Ajouter de petites fixtures pour la valeur de l'ancre, les enfants de débordement, la trace ciblée et les gardes d'erreurs. Vérifier séparément le calcul mutable et le recalcul XLSX source. |
| API Python `recalculate_xlsx_file`/`recalculate_xlsx_bytes`, avec conservation des parties XLSX, contrôle des caches, horloge et graine déterministes | Linexcel analyse les classeurs sans publier un fichier recalculé. Le recalcul amont refuse le classeur complet hors de son sous-ensemble pris en charge. | **P2.** Évaluer un export séparé sur une copie, avec les refus testés et des options de temps/graine explicites. Ne pas l'insérer dans le chemin d'analyse. |
| Rapports de recalcul avec raisons d'erreur, liste complète des fonctions inconnues et tolérance de cache d'une unité au 15e chiffre significatif | Ces règles concernent la nouvelle API de recalcul XLSX; elles ne changent pas automatiquement `Workbook.evaluate_cell` ni la comparaison de cache linexcel. | **P2.** Comparer les reçus Python au rapport de provenance linexcel sur un corpus synthétique avant d'aligner les statuts de cache ou les diagnostics. |

Audit du 6 octobre 2026. Périmètre : passage de 0.9.3 à 0.10.1,
analyse complète, calcul ciblé et exploration des grands classeurs.

## Décision

Adopter les résultats natifs corrects de 0.10.1 et conserver les secours pour
les pannes effectivement rencontrées. Un `#REF!` calculé est une erreur Excel
avec une provenance moteur. Il ne doit pas être transformé en valeur inconnue
ou remplacé par un ancien cache uniquement pour satisfaire un ancien test.

L'intégration possède déjà plusieurs fonctions utiles : évaluation ciblée,
trace bornée des dépendances, import parcimonieux pour les grands fichiers,
écritures groupées des entrées, sorties temporelles en nombres sériels,
probes sémantiques et isolation du moteur. Les prochaines améliorations
doivent prolonger ces mécanismes.

## Échec Dependabot vérifié

La [PR #112](https://github.com/auspect/linexcel/pull/112) met à jour uniquement
le verrou de 0.9.3 vers 0.10.1. La branche locale `codex/formualizer-0.10`
déclare aussi le minimum `formualizer>=0.10.1`.

Le [run CI 37340805701](https://github.com/auspect/linexcel/actions/runs/37340805701)
échoue dans les quatre jobs de tests Linux/Windows. Lint, typage et navigateur
réussissent. Le job Linux/Python 3.13 rapporte 13 échecs, 1335 succès et 18
tests ignorés. La reproduction locale initiale sous Python 3.14 rapporte les
mêmes 13 échecs, 1329 succès et 18 tests ignorés.

Les échecs se répartissent en trois catégories :

| Cause | Conséquence | Correction attendue |
| --- | --- | --- |
| Références de feuille manquante traitées par le moteur | Les tests attendent encore une quarantaine, un secours ou une absence de valeur | Vérifier `#REF!`, le résultat des gardes et la provenance moteur ; injecter une panne pour tester les vrais chemins de secours |
| Formules rendues avec des espaces | Un test de récupération confond le texte rendu et la formule équivalente | Comparer la formule des fixtures sans perdre ses références ; conserver un test effectif de récupération |
| Plans de calcul compressés | Le nombre de couches ne détecte plus certaines longues chaînes | Mesurer les dépendances de la trace avec un parcours borné |

Le test de garde imbriquée est particulièrement utile :
`IFERROR(SUM(IFERROR(NOSHEET!A1,0)),1)` doit produire **0**. Son ancienne
attente, **1**, documentait une limite du secours de linexcel. La nouvelle
valeur native corrige cette limite.

## Nouveautés utiles et décisions d'intégration

Le [changelog amont](https://github.com/PSU3D0/formualizer/blob/main/CHANGELOG.md)
et les [releases](https://github.com/PSU3D0/formualizer/releases) décrivent la
nouvelle autorité de dépendances par régions et les familles de formules.
Cela change la façon d'interpréter les plans, même lorsque les méthodes Python
gardent leurs noms. Les gains publiés par l'amont ne mesurent pas linexcel.

| Possibilité | Situation de linexcel | Prochaine preuve nécessaire |
| --- | --- | --- |
| Familles natives et calcul par blocs | Le moteur est utilisé ; les callbacks et réécritures Python peuvent réduire les chemins natifs disponibles | Profiler un classeur à formules copiées, avec et sans adaptations, après vérification des valeurs et types |
| Gardes d'erreurs et tableaux | Les anciens tests privilégient parfois le secours | Contrats natifs et intégration sur erreurs, tableaux et dépendants |
| Recalcul incrémental | Les tâches interactives reconstruisent leur moteur pour une fermeture choisie | Prototype de session persistante par projet/révision, borné en mémoire, avec isolation et invalidation explicites |
| Écritures groupées | Les entrées littérales sont groupées ; la préparation des étapes écrit encore cellule par cellule | Comparer préparation, calcul et lecture des étapes ; conserver marqueurs, erreurs et règles sur les volatiles |
| Fonctions financières et critères | Des corrections amont peuvent modifier des valeurs auparavant acceptées | Ajouter des fixtures oraculées pour amortissement, tests statistiques et critères échappés selon les usages réels |
| Horloge et hasard contrôlables | La provenance signale déjà les snapshots volatiles | Exposer un contexte de calcul reproductible seulement avec une sémantique claire et le même contexte pour les étapes |
| SheetPort | L'application explore la lignée plutôt qu'un contrat d'entrées/sorties | Tester un cas métier de scénarios typés avant d'ajouter une API publique |

Les stubs Python installés exposent les setters groupés, `trace`,
`inspect_cell`, `evaluate_cells` et les contrôles temporels. Ils n'exposent pas
un itérateur public de formules clairsemées. Une migration vers une telle API
ne peut donc pas être promise sur la seule base des régions internes Rust.

Le wrapper `CompatibleWorkbook.set_formulas_batch` n'a actuellement aucun
appelant. Le setter natif rejette `None`, alors que le wrapper l'utilise pour
préserver les trous. Remplacer ce wrapper seul ne procure donc pas de gain
mesuré et ne constitue pas une optimisation livrée.

Une comparaison locale supplémentaire utilise des entrées numériques sérielles
explicites et une sortie native `serial`, sans les callbacks ni les réécritures
de linexcel. Elle passe 26 des 40 cas, contre 40/40 avec l'intégration. Ce sont
deux configurations de diagnostic, pas deux versions du moteur. Les écarts
concernent des références à du texte contenant un espace, des noms constants
importés et les anciens noms des lois normales. Le chargement peut transformer
le texte en blanc ; ces écarts ne prouvent donc pas seuls une erreur de
comparaison. La suppression d'une adaptation doit tester séparément le
chargement, l'opérateur et la sortie.

Les probes posant directement `10` et `" "` confirment cette différence
d'import : le natif calcule correctement la comparaison et sa garde.
Les 15 formules de comparaison déjà couvertes dans `test_typed_operators`
produisent aussi les mêmes résultats en natif et avec la réécriture. Réduire
`LINEXCEL.COMPAT.COMPARE` devient donc le premier candidat d'optimisation,
après extension de l'oracle aux plages, tableaux, dates et textes Unicode.
Les sorties des tableaux doivent être lues dans leurs cellules de débordement,
car `evaluate_cell` renvoie seulement leur premier élément.

Les adaptations de dates restent nécessaires dans les cas testés :
`DATE(1900,2,29)` produit nativement 61 dans le système 1900 au lieu de 60,
`DAY(60)` produit 28 au lieu de 29, et une date au-delà de l'an 9999 est
acceptée au lieu de rendre `#NUM!`. Les cas de dates sous forme de texte et
de tableaux exigent aussi une validation séparée. Aucun callback de dates
ni d'opérateurs n'a été supprimé dans cette intervention.

## Lots prioritaires

| Priorité | Lot | Critère de livraison |
| --- | --- | --- |
| P0 | Migration 0.10.1 | Suite existante verte ; secours testés par panne injectée ; résultats natifs corrects conservés ; longues chaînes ciblées encore signalées |
| P0 | Contrats et veille CI | Tester version minimale, version verrouillée et dernière stable ; publier version et JUnit ; toute erreur échoue explicitement |
| P1 | Réduire les adaptations Python devenues inutiles | Oracle indépendant pour chaque opérateur/fonction, pour les deux époques, avec erreurs, blancs, textes, tableaux et noms ; ne retirer que l'adaptation prouvée inutile |
| P1 | Préparer les étapes par lots | Mesure avant/après sur petits et grands lots ; mêmes valeurs, types, provenance et échecs ; aucune valeur d'une tentative précédente réutilisée |
| P1 | Comparer les performances de linexcel | Même corpus généré, mêmes budgets et environnement ; temps import/calcul/trace/étapes/export, pic mémoire, couverture et valeurs |
| P2 | Sessions incrémentales interactives | Invalidations après modification d'entrée, formule, structure ou référence externe ; isolation entre projets ; eviction et annulation testées |
| P2 | Contexte reproductible et scénarios SheetPort | Cas utilisateur explicite ; horloge/seed enregistrées ; contrat de types et provenance visible |

Les garde-fous sur la profondeur des formules restent nécessaires. Les
changements du parseur ne suffisent pas à prouver que tous les chemins de
conversion, d'évaluation ou de destruction d'AST sont sûrs.

## À chaque release

1. Lire les notes de la version publiée et les différences depuis la version
   verrouillée. Distinguer les capacités Python publiées des API Rust ou de la
   branche amont non publiée.
2. Exécuter les contrats sur la borne minimale, le verrou et la dernière
   stable. Pour une erreur, produire une petite fixture sans données privées.
3. Comparer les résultats natifs aux oracles et aux adaptations. Une
   correction amont doit pouvoir retirer un secours devenu inutile.
4. Profiler les familles copiées, chaînes, lookups, agrégats conditionnels,
   tableaux dynamiques et classeurs clairsemés. Vérifier les valeurs avant de
   comparer les temps ; exclure un résultat incorrect des gains revendiqués.
5. Choisir un lot utile, avec preuve et borne de coût. Mettre à jour le
   changelog, les limitations et cette matrice, puis lancer les contrôles de
   livraison applicables.

Le workflow récurrent proposé s'active après publication sur la branche par
défaut. Il détecte les incompatibilités ; l'examen des nouveautés et le choix
des améliorations restent une tâche explicite. Il ne fusionne pas les mises
à jour et ne crée pas de messages externes.

Sa sélection exerce les valeurs, dates, calculs ciblés, compatibilités Excel et
calculs sur grands fichiers, ainsi que des contrats financiers/statistiques,
les gardes sur tableaux, les critères échappés et les lecteurs de débordement.
Le rapport d'API capture aussi les membres publics de `Workbook`, `Sheet`,
`EvaluationConfig` et `TraceGraph`, pour repérer les capacités à étudier après
une release. Ces listes ne certifient pas leur sémantique.

## Validation de cette intervention

L'utilisateur a demandé de sauter les contrôles IA pour cette intervention.
La documentation IA, les descriptions d'images et la revue visuelle par IA
sont donc exclues. Les tests de calcul, lint, typage, navigateur et les
captures locales restent dans le périmètre. `validate_manual.py --no-ai
--no-vision` marque volontairement son manifeste `incomplete` et retourne 1 ;
ce statut ne doit pas être présenté comme une acceptation IA complète.

Les preuves locales restent dans `validation_screenshots/`, ignoré par Git.
Les classeurs privés et leurs rapports ne doivent pas être ajoutés au dépôt.

Les deux classeurs générés ont été analysés et recalculés via LibreOffice,
avec 3 captures pour `sales` et 9 pour `stress`. Le manifeste final conserve
le statut `incomplete` attendu sans IA. Le classeur de stress conserve ses
limites explicites : liens externes absents, références non résolues,
itération non convergente et divergences entre moteur et caches LibreOffice.
Les fonctions modernes que LibreOffice conserve en `#NAME?` ne constituent
pas un oracle indépendant utilisable dans ce run.

Le relevé local complet de **1 453 succès et 12 tests ignorés**, les **99 tests
navigateur**, les 37 tests ciblés et les deux distributions construites
proviennent de l'arbre avant le rebase sur linexcel 1.13.0. Le manifeste local
de ce run garde le statut `incomplete` demandé, puisqu'il omet les contrôles IA.
Ces nombres et artefacts décrivent cet arbre; ils ne valident pas 0.11.0.

Depuis, le workflow a tourné dans GitHub : le job navigateur passe dans le
[run PR #113](https://github.com/auspect/linexcel/actions/runs/37619156812),
et les versions plancher/verrou passent **236 tests** dans le
[run 37619156894](https://github.com/auspect/linexcel/actions/runs/37619156894).
Le premier job latest sur formualizer 0.11.0 a rapporté **116 échecs et 120 succès** à
cause des collisions entre fonctions statistiques natives et alias linexcel.
Les correctifs décrits dans l'addendum répondent à ces collisions et aux
limites de complexité du parseur. La sélection finale de **27 tests** de garde
et de robustesse passe sur **0.10.1 et 0.11.0**. Ruff, formatage et typage
passent aussi. Les résultats GitHub de la branche publiée sont consultables
dans les [contrôles de la PR #113](https://github.com/auspect/linexcel/pull/113/checks).
Le tag de release reste conditionné à leur réussite.
