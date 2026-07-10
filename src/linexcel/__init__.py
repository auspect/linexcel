"""Analyse de data lineage pour classeurs Excel.

Pipeline : extraction des formules (formualizer, moteur Rust) →
regroupement des formules étirées (canonicalisation R1C1) →
graphe de dépendances (cellules, plages, noms définis, VBA) →
décomposition des fonctions composées avec évaluation pas-à-pas →
documentation automatique par IA (google-genai) fondée sur le
lignage déterministe.

Utilisable comme bibliothèque autonome (marimo, Jupyter, script), sans
serveur FastAPI ni clé IA :

    from linexcel import analyze
    result = analyze("classeur.xlsx")   # -> LineageResult
    result                               # graphe interactif dans marimo
    result.save_html("lineage.html")     # visualiseur autonome
    result.stats, result.warnings        # métadonnées
    result.document(api_key="…")         # documentation IA (optionnelle)
"""

from linexcel.analyzer import analyze_workbook
from linexcel.result import LineageResult, analyze

__all__ = ["analyze", "LineageResult", "analyze_workbook"]
