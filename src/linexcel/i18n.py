"""Supported languages for the viewer UI and the AI documentation.

The set is deliberately **closed**. ``language`` selects a stored system prompt
and is interpolated into the generated viewer, so it is validated against this
allowlist rather than escaped or passed through: an arbitrary string would let a
caller steer the model's instructions or reach the generated JavaScript.

Adding a language means adding an entry here *and* in :mod:`linexcel.aidoc`'s
two prompt registries. ``tests/test_lineage.py`` asserts the three stay in sync,
so a partial addition fails the suite instead of surfacing as raw keys in the
report or a ``KeyError`` at generation time.

Provenance: ``en`` and ``fr`` were written by hand. The other seven languages,
here and in the prompt registries, were produced with AI assistance and have not
been reviewed by native speakers — corrections welcome.
"""

from __future__ import annotations

from typing import Any

#: Languages accepted by ``language=`` across the public API.
LANGUAGES = ("en", "fr", "es", "de", "it", "pt", "nl", "ja", "zh")

DEFAULT_LANGUAGE = "en"

#: Viewer interface strings. Every language carries the same key set; ``{name}``
#: placeholders are substituted client-side and must be preserved verbatim.
UI_STRINGS: dict[str, dict[str, str]] = {
    "en": {
        "graph": "Graph",
        "overview": "Workbook overview",
        "visual": "Visual preview",
        "search": "Search… ⏎",
        "search_matches": "{count} matches",
        "fit_all": "Fit",
        "fit_sel": "Fit Selection",
        "zoom_in": "Zoom In",
        "zoom_out": "Zoom Out",
        "formula": "Formula",
        "stretched_pattern": "Stretched pattern over {count} cells ({bbox}).",
        "computed_value": "Computed value",
        "value_samples": "Value samples",
        "value_from_file": "Read from Excel file",
        "value_recalc": "Recalculated by linexcel",
        "value_fallback": "linexcel fallback (error-guarded)",
        "value_cached": "Excel file value",
        "differs_from_file": "{recalc} differs from the file's value",
        "target": "Target",
        "step_decomp": "Step-by-step decomposition",
        "step_hint": "Each function/operator is evaluated individually.",
        "final_result": "Final result",
        "not_evaluated": "not evaluated",
        "precedents": "Precedents",
        "dependents": "Dependents",
        "cells": "cells",
        "ai_doc": "🤖 AI Documentation (Generated)",
        "ai_overview": "🤖 AI Generated Overview",
        "ai_overview_desc": (
            "This overview was written by an AI model from the deterministic "
            "lineage. The facts presented are derived from the workbook's "
            "formulas and data."
        ),
        "fallback": (
            "Cytoscape could not be loaded (CDN access required). The JSON "
            "graph remains available via result.to_dict()."
        ),
        "stats": "{formulas} formulas · {nodes} nodes · {edges} edges{vba}",
        "kind_cell": "Formula",
        "kind_group": "Stretched formulas",
        "kind_input": "Source data",
        "kind_name": "Named cell/range",
        "kind_vba": "VBA",
        "kind_misc": "Other (aggregated)",
        "kind_opaque": "External reference",
        "placeholder_title": "Select a node",
        "placeholder_desc": (
            "Select a node in the graph to inspect its formula, computed "
            "value, step-by-step evaluation, and AI-generated documentation."
        ),
        "sheets_tab": "Sheets",
        "all_sheets": "All sheets",
        "sheet_filter": "Sheet filter",
        "page": "Page {n}",
    },
    "fr": {
        "graph": "Graphe",
        "overview": "Synthèse générale",
        "visual": "Aperçu visuel",
        "search": "Rechercher… ⏎",
        "search_matches": "{count} résultats",
        "fit_all": "Ajuster",
        "fit_sel": "Ajuster la sélection",
        "zoom_in": "Zoom avant",
        "zoom_out": "Zoom arrière",
        "formula": "Formule",
        "stretched_pattern": "Formule étirée sur {count} cellules ({bbox}).",
        "computed_value": "Valeur calculée",
        "value_samples": "Échantillons de valeurs",
        "value_from_file": "Lue du fichier Excel",
        "value_recalc": "Recalculée par linexcel",
        "value_fallback": "repli linexcel (erreur protégée)",
        "value_cached": "Valeur du fichier Excel",
        "differs_from_file": "{recalc} diffère de la valeur du fichier",
        "target": "Cible",
        "step_decomp": "Décomposition pas-à-pas",
        "step_hint": "Chaque fonction et opérateur est évalué individuellement.",
        "final_result": "Résultat final",
        "not_evaluated": "non évalué",
        "precedents": "Précédents",
        "dependents": "Dépendants",
        "cells": "cellules",
        "ai_doc": "🤖 Documentation IA (Générée)",
        "ai_overview": "🤖 Synthèse Générée par IA",
        "ai_overview_desc": (
            "Cette synthèse a été rédigée par un modèle d'IA à partir du "
            "lignage de calculs déterministe. Les faits présentés proviennent "
            "des formules et des données du classeur."
        ),
        "fallback": (
            "Cytoscape n'a pas pu être chargé (accès CDN requis). Le graphe "
            "JSON reste disponible via result.to_dict()."
        ),
        "stats": "{formulas} formules · {nodes} nœuds · {edges} liens{vba}",
        "kind_cell": "Formule",
        "kind_group": "Formules étirées",
        "kind_input": "Source de données",
        "kind_name": "Cellule/plage nommée",
        "kind_vba": "VBA",
        "kind_misc": "Autre (agrégé)",
        "kind_opaque": "Référence externe",
        "placeholder_title": "Sélectionner un nœud",
        "placeholder_desc": (
            "Sélectionnez un nœud dans le graphe pour afficher sa formule, sa "
            "valeur calculée, sa décomposition pas à pas et sa documentation IA."
        ),
        "sheets_tab": "Feuilles",
        "all_sheets": "Toutes les feuilles",
        "sheet_filter": "Filtre par feuille",
        "page": "Page {n}",
    },
    "es": {
        "graph": "Grafo",
        "overview": "Resumen del libro",
        "visual": "Vista previa visual",
        "search": "Buscar… ⏎",
        "search_matches": "{count} resultados",
        "fit_all": "Ajustar",
        "fit_sel": "Ajustar selección",
        "zoom_in": "Acercar",
        "zoom_out": "Alejar",
        "formula": "Fórmula",
        "stretched_pattern": "Patrón extendido sobre {count} celdas ({bbox}).",
        "computed_value": "Valor calculado",
        "value_samples": "Muestras de valores",
        "value_from_file": "Leído del archivo Excel",
        "value_recalc": "Recalculado por linexcel",
        "value_fallback": "valor alternativo de linexcel (error protegido)",
        "value_cached": "Valor del archivo Excel",
        "differs_from_file": "{recalc} difiere del valor del archivo",
        "target": "Destino",
        "step_decomp": "Descomposición paso a paso",
        "step_hint": "Cada función y operador se evalúa individualmente.",
        "final_result": "Resultado final",
        "not_evaluated": "no evaluado",
        "precedents": "Precedentes",
        "dependents": "Dependientes",
        "cells": "celdas",
        "ai_doc": "🤖 Documentación IA (generada)",
        "ai_overview": "🤖 Resumen generado por IA",
        "ai_overview_desc": (
            "Este resumen ha sido redactado por un modelo de IA a partir del "
            "linaje de cálculos determinista. Los hechos presentados proceden "
            "de las fórmulas y los datos del libro."
        ),
        "fallback": (
            "No se pudo cargar Cytoscape (se requiere acceso al CDN). El grafo "
            "JSON sigue disponible mediante result.to_dict()."
        ),
        "stats": "{formulas} fórmulas · {nodes} nodos · {edges} enlaces{vba}",
        "kind_cell": "Fórmula",
        "kind_group": "Fórmulas extendidas",
        "kind_input": "Datos de origen",
        "kind_name": "Celda/rango con nombre",
        "kind_vba": "VBA",
        "kind_misc": "Otros (agregados)",
        "kind_opaque": "Referencia externa",
        "placeholder_title": "Seleccione un nodo",
        "placeholder_desc": (
            "Seleccione un nodo del grafo para consultar su fórmula, su valor "
            "calculado, su evaluación paso a paso y su documentación generada "
            "por IA."
        ),
        "sheets_tab": "Hojas",
        "all_sheets": "Todas las hojas",
        "sheet_filter": "Filtro por hoja",
        "page": "Página {n}",
    },
    "de": {
        "graph": "Graph",
        "overview": "Arbeitsmappen-Überblick",
        "visual": "Visuelle Vorschau",
        "search": "Suchen… ⏎",
        "search_matches": "{count} Treffer",
        "fit_all": "Einpassen",
        "fit_sel": "Auswahl einpassen",
        "zoom_in": "Vergrößern",
        "zoom_out": "Verkleinern",
        "formula": "Formel",
        "stretched_pattern": "Gezogenes Muster über {count} Zellen ({bbox}).",
        "computed_value": "Berechneter Wert",
        "value_samples": "Wertebeispiele",
        "value_from_file": "Aus der Excel-Datei gelesen",
        "value_recalc": "Von linexcel neu berechnet",
        "value_fallback": "linexcel-Ersatzwert (fehlergeschützt)",
        "value_cached": "Wert aus der Excel-Datei",
        "differs_from_file": "{recalc} weicht vom Wert der Datei ab",
        "target": "Ziel",
        "step_decomp": "Schrittweise Zerlegung",
        "step_hint": "Jede Funktion und jeder Operator wird einzeln ausgewertet.",
        "final_result": "Endergebnis",
        "not_evaluated": "nicht ausgewertet",
        "precedents": "Vorgänger",
        "dependents": "Nachfolger",
        "cells": "Zellen",
        "ai_doc": "🤖 KI-Dokumentation (generiert)",
        "ai_overview": "🤖 KI-generierter Überblick",
        "ai_overview_desc": (
            "Dieser Überblick wurde von einem KI-Modell auf Basis der "
            "deterministischen Berechnungsherkunft verfasst. Die dargestellten "
            "Fakten stammen aus den Formeln und Daten der Arbeitsmappe."
        ),
        "fallback": (
            "Cytoscape konnte nicht geladen werden (CDN-Zugriff erforderlich). "
            "Der JSON-Graph bleibt über result.to_dict() verfügbar."
        ),
        "stats": "{formulas} Formeln · {nodes} Knoten · {edges} Kanten{vba}",
        "kind_cell": "Formel",
        "kind_group": "Gezogene Formeln",
        "kind_input": "Quelldaten",
        "kind_name": "Benannte Zelle/Bereich",
        "kind_vba": "VBA",
        "kind_misc": "Sonstige (aggregiert)",
        "kind_opaque": "Externer Bezug",
        "placeholder_title": "Knoten auswählen",
        "placeholder_desc": (
            "Wählen Sie einen Knoten im Graphen aus, um seine Formel, seinen "
            "berechneten Wert, seine schrittweise Auswertung und seine "
            "KI-generierte Dokumentation anzuzeigen."
        ),
        "sheets_tab": "Blätter",
        "all_sheets": "Alle Blätter",
        "sheet_filter": "Blattfilter",
        "page": "Seite {n}",
    },
    "it": {
        "graph": "Grafo",
        "overview": "Panoramica della cartella",
        "visual": "Anteprima visiva",
        "search": "Cerca… ⏎",
        "search_matches": "{count} risultati",
        "fit_all": "Adatta",
        "fit_sel": "Adatta selezione",
        "zoom_in": "Ingrandisci",
        "zoom_out": "Riduci",
        "formula": "Formula",
        "stretched_pattern": "Motivo esteso su {count} celle ({bbox}).",
        "computed_value": "Valore calcolato",
        "value_samples": "Campioni di valori",
        "value_from_file": "Letto dal file Excel",
        "value_recalc": "Ricalcolato da linexcel",
        "value_fallback": "valore di ripiego di linexcel (errore protetto)",
        "value_cached": "Valore del file Excel",
        "differs_from_file": "{recalc} differisce dal valore del file",
        "target": "Destinazione",
        "step_decomp": "Scomposizione passo passo",
        "step_hint": "Ogni funzione e operatore viene valutato singolarmente.",
        "final_result": "Risultato finale",
        "not_evaluated": "non valutato",
        "precedents": "Precedenti",
        "dependents": "Dipendenti",
        "cells": "celle",
        "ai_doc": "🤖 Documentazione IA (generata)",
        "ai_overview": "🤖 Panoramica generata dall'IA",
        "ai_overview_desc": (
            "Questa panoramica è stata redatta da un modello di IA a partire "
            "dalla derivazione deterministica dei calcoli. I fatti presentati "
            "provengono dalle formule e dai dati della cartella di lavoro."
        ),
        "fallback": (
            "Impossibile caricare Cytoscape (è necessario l'accesso al CDN). "
            "Il grafo JSON resta disponibile tramite result.to_dict()."
        ),
        "stats": "{formulas} formule · {nodes} nodi · {edges} archi{vba}",
        "kind_cell": "Formula",
        "kind_group": "Formule estese",
        "kind_input": "Dati di origine",
        "kind_name": "Cella/intervallo denominato",
        "kind_vba": "VBA",
        "kind_misc": "Altro (aggregato)",
        "kind_opaque": "Riferimento esterno",
        "placeholder_title": "Seleziona un nodo",
        "placeholder_desc": (
            "Seleziona un nodo nel grafo per consultarne la formula, il valore "
            "calcolato, la valutazione passo passo e la documentazione "
            "generata dall'IA."
        ),
        "sheets_tab": "Fogli",
        "all_sheets": "Tutti i fogli",
        "sheet_filter": "Filtro per foglio",
        "page": "Pagina {n}",
    },
    "pt": {
        "graph": "Grafo",
        "overview": "Visão geral da pasta",
        "visual": "Pré-visualização",
        "search": "Pesquisar… ⏎",
        "search_matches": "{count} resultados",
        "fit_all": "Ajustar",
        "fit_sel": "Ajustar seleção",
        "zoom_in": "Ampliar",
        "zoom_out": "Reduzir",
        "formula": "Fórmula",
        "stretched_pattern": "Padrão estendido por {count} células ({bbox}).",
        "computed_value": "Valor calculado",
        "value_samples": "Amostras de valores",
        "value_from_file": "Lido do ficheiro Excel",
        "value_recalc": "Recalculado pelo linexcel",
        "value_fallback": "valor alternativo do linexcel (erro protegido)",
        "value_cached": "Valor do ficheiro Excel",
        "differs_from_file": "{recalc} difere do valor do ficheiro",
        "target": "Destino",
        "step_decomp": "Decomposição passo a passo",
        "step_hint": "Cada função e operador é avaliado individualmente.",
        "final_result": "Resultado final",
        "not_evaluated": "não avaliado",
        "precedents": "Precedentes",
        "dependents": "Dependentes",
        "cells": "células",
        "ai_doc": "🤖 Documentação de IA (gerada)",
        "ai_overview": "🤖 Visão geral gerada por IA",
        "ai_overview_desc": (
            "Esta visão geral foi redigida por um modelo de IA a partir da "
            "linhagem determinista dos cálculos. As informações apresentadas "
            "provêm das fórmulas e dos dados da pasta de trabalho."
        ),
        "fallback": (
            "Não foi possível carregar o Cytoscape (é necessário acesso ao "
            "CDN). O grafo JSON continua disponível através de result.to_dict()."
        ),
        "stats": "{formulas} fórmulas · {nodes} nós · {edges} ligações{vba}",
        "kind_cell": "Fórmula",
        "kind_group": "Fórmulas estendidas",
        "kind_input": "Dados de origem",
        "kind_name": "Célula/intervalo nomeado",
        "kind_vba": "VBA",
        "kind_misc": "Outros (agregados)",
        "kind_opaque": "Referência externa",
        "placeholder_title": "Selecione um nó",
        "placeholder_desc": (
            "Selecione um nó no grafo para consultar a sua fórmula, o valor "
            "calculado, a avaliação passo a passo e a documentação gerada por "
            "IA."
        ),
        "sheets_tab": "Folhas",
        "all_sheets": "Todas as folhas",
        "sheet_filter": "Filtro por folha",
        "page": "Página {n}",
    },
    "nl": {
        "graph": "Graaf",
        "overview": "Werkmapoverzicht",
        "visual": "Visuele weergave",
        "search": "Zoeken… ⏎",
        "search_matches": "{count} resultaten",
        "fit_all": "Passend maken",
        "fit_sel": "Selectie passend maken",
        "zoom_in": "Inzoomen",
        "zoom_out": "Uitzoomen",
        "formula": "Formule",
        "stretched_pattern": "Doorgetrokken patroon over {count} cellen ({bbox}).",
        "computed_value": "Berekende waarde",
        "value_samples": "Voorbeeldwaarden",
        "value_from_file": "Gelezen uit het Excel-bestand",
        "value_recalc": "Herberekend door linexcel",
        "value_fallback": "terugvalwaarde van linexcel (foutbeveiligd)",
        "value_cached": "Waarde uit het Excel-bestand",
        "differs_from_file": "{recalc} wijkt af van de waarde in het bestand",
        "target": "Doel",
        "step_decomp": "Stapsgewijze ontleding",
        "step_hint": "Elke functie en operator wordt afzonderlijk geëvalueerd.",
        "final_result": "Eindresultaat",
        "not_evaluated": "niet geëvalueerd",
        "precedents": "Voorgangers",
        "dependents": "Afhankelijken",
        "cells": "cellen",
        "ai_doc": "🤖 AI-documentatie (gegenereerd)",
        "ai_overview": "🤖 Door AI gegenereerd overzicht",
        "ai_overview_desc": (
            "Dit overzicht is geschreven door een AI-model op basis van de "
            "deterministische herkomst van de berekeningen. De weergegeven "
            "gegevens komen uit de formules en gegevens van de werkmap."
        ),
        "fallback": (
            "Cytoscape kon niet worden geladen (CDN-toegang vereist). De "
            "JSON-graaf blijft beschikbaar via result.to_dict()."
        ),
        "stats": "{formulas} formules · {nodes} knopen · {edges} verbindingen{vba}",
        "kind_cell": "Formule",
        "kind_group": "Doorgetrokken formules",
        "kind_input": "Brongegevens",
        "kind_name": "Benoemde cel/bereik",
        "kind_vba": "VBA",
        "kind_misc": "Overig (samengevoegd)",
        "kind_opaque": "Externe verwijzing",
        "placeholder_title": "Selecteer een knoop",
        "placeholder_desc": (
            "Selecteer een knoop in de graaf om de formule, de berekende "
            "waarde, de stapsgewijze evaluatie en de door AI gegenereerde "
            "documentatie te bekijken."
        ),
        "sheets_tab": "Bladen",
        "all_sheets": "Alle bladen",
        "sheet_filter": "Bladfilter",
        "page": "Pagina {n}",
    },
    "ja": {
        "graph": "グラフ",
        "overview": "ブック概要",
        "visual": "ビジュアルプレビュー",
        "search": "検索… ⏎",
        "search_matches": "{count} 件",
        "fit_all": "全体表示",
        "fit_sel": "選択範囲を表示",
        "zoom_in": "拡大",
        "zoom_out": "縮小",
        "formula": "数式",
        "stretched_pattern": "{count} セル（{bbox}）にコピーされた数式パターンです。",
        "computed_value": "計算結果",
        "value_samples": "値のサンプル",
        "value_from_file": "Excel ファイルから読み取り",
        "value_recalc": "linexcel が再計算",
        "value_fallback": "linexcel の代替値（エラー保護）",
        "value_cached": "Excel ファイルの値",
        "differs_from_file": "{recalc} はファイルの値と異なります",
        "target": "対象範囲",
        "step_decomp": "ステップごとの分解",
        "step_hint": "各関数・演算子を個別に評価しています。",
        "final_result": "最終結果",
        "not_evaluated": "未評価",
        "precedents": "参照元",
        "dependents": "参照先",
        "cells": "セル",
        "ai_doc": "🤖 AI ドキュメント（生成）",
        "ai_overview": "🤖 AI が生成した概要",
        "ai_overview_desc": (
            "この概要は、決定論的に抽出された計算系統に基づいて AI モデルが"
            "作成したものです。記載内容はブックの数式とデータに由来します。"
        ),
        "fallback": (
            "Cytoscape を読み込めませんでした（CDN へのアクセスが必要です）。"
            "JSON グラフは result.to_dict() から引き続き利用できます。"
        ),
        "stats": "数式 {formulas} 件 · ノード {nodes} 件 · エッジ {edges} 件{vba}",
        "kind_cell": "数式",
        "kind_group": "コピーされた数式",
        "kind_input": "ソースデータ",
        "kind_name": "名前付きセル/範囲",
        "kind_vba": "VBA",
        "kind_misc": "その他（集約）",
        "kind_opaque": "外部参照",
        "placeholder_title": "ノードを選択してください",
        "placeholder_desc": (
            "グラフ内のノードを選択すると、数式、計算結果、ステップごとの"
            "評価、AI が生成したドキュメントを確認できます。"
        ),
        "sheets_tab": "シート",
        "all_sheets": "すべてのシート",
        "sheet_filter": "シートで絞り込み",
        "page": "ページ {n}",
    },
    "zh": {
        "graph": "图谱",
        "overview": "工作簿概览",
        "visual": "可视化预览",
        "search": "搜索… ⏎",
        "search_matches": "{count} 个匹配",
        "fit_all": "适应窗口",
        "fit_sel": "适应所选",
        "zoom_in": "放大",
        "zoom_out": "缩小",
        "formula": "公式",
        "stretched_pattern": "填充公式模式，覆盖 {count} 个单元格（{bbox}）。",
        "computed_value": "计算值",
        "value_samples": "取值示例",
        "value_from_file": "读取自 Excel 文件",
        "value_recalc": "由 linexcel 重新计算",
        "value_fallback": "linexcel 回退值（已防错）",
        "value_cached": "Excel 文件中的值",
        "differs_from_file": "{recalc} 与文件中的值不一致",
        "target": "目标区域",
        "step_decomp": "逐步分解",
        "step_hint": "每个函数和运算符均单独求值。",
        "final_result": "最终结果",
        "not_evaluated": "未求值",
        "precedents": "引用单元格",
        "dependents": "从属单元格",
        "cells": "个单元格",
        "ai_doc": "🤖 AI 文档（自动生成）",
        "ai_overview": "🤖 AI 生成的概览",
        "ai_overview_desc": (
            "本概览由 AI 模型根据确定性的计算血缘生成。"
            "所述内容均来自工作簿的公式与数据。"
        ),
        "fallback": (
            "无法加载 Cytoscape（需要访问 CDN）。"
            "JSON 图谱仍可通过 result.to_dict() 获取。"
        ),
        "stats": "{formulas} 个公式 · {nodes} 个节点 · {edges} 条边{vba}",
        "kind_cell": "公式",
        "kind_group": "填充公式",
        "kind_input": "源数据",
        "kind_name": "命名单元格/区域",
        "kind_vba": "VBA",
        "kind_misc": "其他（已聚合）",
        "kind_opaque": "外部引用",
        "placeholder_title": "请选择一个节点",
        "placeholder_desc": (
            "在图谱中选择一个节点，即可查看其公式、计算值、逐步求值过程"
            "以及 AI 生成的文档。"
        ),
        "sheets_tab": "工作表",
        "all_sheets": "所有工作表",
        "sheet_filter": "按工作表筛选",
        "page": "第 {n} 页",
    },
}


def validate_language(language: str) -> str:
    """Return ``language`` if it is supported, else raise ``ValueError``.

    >>> validate_language("ja")
    'ja'
    >>> validate_language("klingon")
    Traceback (most recent call last):
        ...
    ValueError: Unsupported language: 'klingon'. Use one of
    ('en', 'fr', 'es', 'de', 'it', 'pt', 'nl', 'ja', 'zh')
    """
    if language not in LANGUAGES:
        raise ValueError(f"Unsupported language: {language!r}. Use one of {LANGUAGES}")
    return language


def ui_payload(language: str) -> dict[str, Any]:
    """Interface strings to embed for ``language``.

    Only the requested language and the English fallback are shipped: the
    viewer resolves a missing key through ``I18N.en``, so embedding the other
    seven locales in every report would be dead weight.
    """
    validate_language(language)
    payload: dict[str, Any] = {DEFAULT_LANGUAGE: UI_STRINGS[DEFAULT_LANGUAGE]}
    payload[language] = UI_STRINGS[language]
    return payload
