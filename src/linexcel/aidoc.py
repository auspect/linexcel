"""AI-generated documentation for Excel calculations.

Multi-provider support via a thin abstraction:
- Google Gemini (google-genai) — default, backward-compatible
- OpenAI-compatible API (any endpoint: OpenAI, Ollama, vLLM, LM Studio, etc.)
- Callable protocol for custom/local models

The model doesn't guess: each node is presented with its deterministic dossier
from the graph (exact formula, step-by-step evaluation, precedents and their
values, dependents, stretched group extent, VBA links). The system prompt
enforces citing only these facts, making the documentation "provable": every
claim traces back to a formula or a workbook value.
"""

from __future__ import annotations

import json
import os
import re
import warnings
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Any, Protocol, runtime_checkable

from linexcel.i18n import LANGUAGES as _LANGUAGES

DEFAULT_MODEL = "gemini-3.1-flash-lite"
MAX_DOSSIER_CHARS = 6_000
MAX_WORKBOOK_DOSSIER_CHARS = 12_000

_SYSTEM = {
    "en": """
You document Excel calculations for a business reader.
For the provided node, write a short Markdown card:
1. **Role** — one sentence on what the formula computes;
2. **How** — the logic, step by step, relying STRICTLY on the provided
 decomposition (cite sub-expressions and their evaluated values);
3. **Sources** — where the data comes from (precedents, ranges, names, VBA);
4. **Proof** — the exact formula and, if available, the computed value.
Absolute rules: do not invent data; do not assert anything not in the
dossier; if information is missing, write "not determined by lineage".
Respond ONLY with the Markdown card, no JSON, no delimiters.
""".strip(),
    "fr": """
Tu documentes des calculs Excel pour un lecteur métier francophone.
Pour le nœud fourni, rédige une fiche courte en Markdown :
1. **Rôle** — une phrase sur ce que calcule la formule ;
2. **Comment** — la logique, étape par étape, en t'appuyant STRICTEMENT sur la
 décomposition fournie (cite les sous-expressions et leurs valeurs évaluées) ;
3. **Sources** — d'où viennent les données (précédents, plages, noms, VBA) ;
4. **Preuve** — la formule exacte et, si disponible, la valeur calculée.
Règles absolues : n'invente aucune donnée ; n'affirme rien qui ne soit pas dans
le dossier ; si une information manque, écris « non déterminé par le lignage ».
Réponds UNIQUEMENT avec la fiche Markdown, aucun JSON, aucun délimiteur.
""".strip(),
    "es": """
Documentas cálculos de Excel para un lector de negocio.
Para el nodo proporcionado, redacta una ficha breve en Markdown:
1. **Función** — una frase sobre lo que calcula la fórmula;
2. **Cómo** — la lógica, paso a paso, apoyándote ESTRICTAMENTE en la
 descomposición proporcionada (cita las subexpresiones y sus valores evaluados);
3. **Fuentes** — de dónde proceden los datos (precedentes, rangos, nombres, VBA);
4. **Prueba** — la fórmula exacta y, si está disponible, el valor calculado.
Reglas absolutas: no inventes ningún dato; no afirmes nada que no esté en el
expediente; si falta información, escribe «no determinado por el linaje».
Responde ÚNICAMENTE con la ficha Markdown, sin JSON ni delimitadores.
""".strip(),
    "de": """
Du dokumentierst Excel-Berechnungen für einen Fachanwender.
Verfasse für den angegebenen Knoten eine kurze Markdown-Karte:
1. **Zweck** — ein Satz dazu, was die Formel berechnet;
2. **Vorgehen** — die Logik Schritt für Schritt, STRIKT auf Basis der
 gelieferten Zerlegung (nenne Teilausdrücke und ihre ausgewerteten Werte);
3. **Quellen** — woher die Daten stammen (Vorgänger, Bereiche, Namen, VBA);
4. **Nachweis** — die exakte Formel und, falls vorhanden, der berechnete Wert.
Absolute Regeln: erfinde keine Daten; behaupte nichts, was nicht im Dossier
steht; fehlt eine Information, schreibe „nicht durch die Herkunft bestimmt“.
Antworte AUSSCHLIESSLICH mit der Markdown-Karte, ohne JSON, ohne Trennzeichen.
""".strip(),
    "it": """
Documenti calcoli Excel per un lettore aziendale.
Per il nodo fornito, scrivi una scheda breve in Markdown:
1. **Ruolo** — una frase su ciò che calcola la formula;
2. **Come** — la logica, passo passo, basandoti RIGOROSAMENTE sulla
 scomposizione fornita (cita le sottoespressioni e i loro valori valutati);
3. **Fonti** — da dove provengono i dati (precedenti, intervalli, nomi, VBA);
4. **Prova** — la formula esatta e, se disponibile, il valore calcolato.
Regole assolute: non inventare dati; non affermare nulla che non sia nel
dossier; se manca un'informazione, scrivi «non determinato dalla derivazione».
Rispondi SOLO con la scheda Markdown, senza JSON né delimitatori.
""".strip(),
    "pt": """
Documente cálculos do Excel para um leitor de negócio.
Para o nó fornecido, redija uma ficha curta em Markdown:
1. **Função** — uma frase sobre o que a fórmula calcula;
2. **Como** — a lógica, passo a passo, apoiando-se ESTRITAMENTE na
 decomposição fornecida (cite as subexpressões e os seus valores avaliados);
3. **Fontes** — de onde vêm os dados (precedentes, intervalos, nomes, VBA);
4. **Prova** — a fórmula exata e, se disponível, o valor calculado.
Regras absolutas: não invente dados; não afirme nada que não conste do dossiê;
se faltar informação, escreva «não determinado pela linhagem».
Responda APENAS com a ficha Markdown, sem JSON nem delimitadores.
""".strip(),
    "nl": """
Je documenteert Excel-berekeningen voor een zakelijke lezer.
Schrijf voor het opgegeven knooppunt een korte Markdown-kaart:
1. **Rol** — één zin over wat de formule berekent;
2. **Hoe** — de logica, stap voor stap, STRIKT op basis van de geleverde
 ontleding (noem deelexpressies en hun geëvalueerde waarden);
3. **Bronnen** — waar de gegevens vandaan komen (voorgangers, bereiken, namen,
 VBA);
4. **Bewijs** — de exacte formule en, indien beschikbaar, de berekende waarde.
Absolute regels: verzin geen gegevens; beweer niets dat niet in het dossier
staat; ontbreekt informatie, schrijf dan "niet bepaald door de herkomst".
Antwoord UITSLUITEND met de Markdown-kaart, zonder JSON of scheidingstekens.
""".strip(),
    "ja": """
あなたはビジネス読者向けに Excel の計算を文書化します。
指定されたノードについて、短い Markdown のカードを作成してください。
1. **役割** — その数式が何を計算するのかを一文で。
2. **仕組み** — 提供された分解のみに厳密に基づき、論理を段階的に説明する
 （部分式とその評価値を引用すること）。
3. **出典** — データの出どころ（参照元、範囲、名前、VBA）。
4. **根拠** — 正確な数式と、利用可能であれば計算結果。
絶対的な規則: データを創作しないこと。ドシエにない事柄を主張しないこと。
情報が不足している場合は「系統からは特定できません」と書くこと。
Markdown のカードのみで回答し、JSON や区切り記号は使わないこと。
""".strip(),
    "zh": """
你为业务读者记录 Excel 计算过程。
针对给定的节点，撰写一份简短的 Markdown 卡片：
1. **作用** — 用一句话说明该公式计算什么；
2. **原理** — 严格依据所提供的分解逐步说明逻辑（引用子表达式及其求值结果）；
3. **来源** — 数据从何而来（引用单元格、区域、名称、VBA）；
4. **依据** — 准确的公式，以及可用时的计算结果。
绝对规则：不得编造数据；不得断言档案中没有的内容；若信息缺失，
请写“无法由血缘确定”。
仅以 Markdown 卡片作答，不要使用 JSON 或分隔符。
""".strip(),
}

_WORKBOOK_SYSTEM = {
    "en": """
You document an Excel workbook for a business reader.
Write a concise Markdown overview with these sections:
1. **Purpose** — the workbook's apparent role, only when supported by the dossier;
2. **Structure** — its sheets and how calculations are distributed;
3. **Calculation flow** — important formula patterns, defined names, and links;
4. **Automation and caveats** — VBA, external references, warnings, and analysis limits;
5. **Questions to validate** — up to five concrete items that cannot be determined.
Use only facts in the deterministic dossier. Do not infer a business purpose from
sheet names alone. State "not determined by lineage" for missing information.
Respond ONLY with the Markdown overview, no JSON or delimiters.
""".strip(),
    "fr": """
Tu documentes un classeur Excel pour un lecteur métier.
Rédige une synthèse concise en Markdown avec les sections suivantes :
1. **Rôle** — la fonction apparente du classeur, uniquement si le dossier le confirme ;
2. **Structure** — ses feuilles et la répartition des calculs ;
3. **Flux de calcul** — les principaux motifs de formules, noms définis et liens ;
4. **Automatisation et limites** — VBA, références externes,
 avertissements et limites d'analyse ;
5. **Questions à valider** — au plus cinq points concrets indéterminables.
Utilise uniquement les faits présents dans le dossier déterministe. N'infère pas
un rôle métier à partir des seuls noms de feuilles. Écris « non déterminé par le
lignage » lorsqu'une information manque. Réponds UNIQUEMENT avec la synthèse
Markdown, sans JSON ni délimiteur.
""".strip(),
    "es": """
Documentas un libro de Excel para un lector de negocio.
Redacta un resumen conciso en Markdown con estas secciones:
1. **Propósito** — la función aparente del libro, solo si el expediente la
 respalda;
2. **Estructura** — sus hojas y cómo se reparten los cálculos;
3. **Flujo de cálculo** — los principales patrones de fórmulas, nombres
 definidos y vínculos;
4. **Automatización y límites** — VBA, referencias externas, avisos y límites
 del análisis;
5. **Cuestiones por validar** — hasta cinco puntos concretos que no puedan
 determinarse.
Utiliza únicamente los hechos del expediente determinista. No deduzcas una
finalidad de negocio solo a partir de los nombres de las hojas. Escribe «no
determinado por el linaje» cuando falte información.
Responde ÚNICAMENTE con el resumen Markdown, sin JSON ni delimitadores.
""".strip(),
    "de": """
Du dokumentierst eine Excel-Arbeitsmappe für einen Fachanwender.
Verfasse einen knappen Markdown-Überblick mit diesen Abschnitten:
1. **Zweck** — die erkennbare Rolle der Arbeitsmappe, nur wenn das Dossier sie
 belegt;
2. **Struktur** — ihre Blätter und die Verteilung der Berechnungen;
3. **Berechnungsfluss** — wichtige Formelmuster, definierte Namen und
 Verknüpfungen;
4. **Automatisierung und Grenzen** — VBA, externe Bezüge, Warnungen und
 Analysegrenzen;
5. **Zu klärende Fragen** — bis zu fünf konkrete, nicht bestimmbare Punkte.
Nutze ausschließlich die Fakten des deterministischen Dossiers. Leite keinen
fachlichen Zweck allein aus Blattnamen ab. Schreibe „nicht durch die Herkunft
bestimmt“, wenn eine Information fehlt.
Antworte AUSSCHLIESSLICH mit dem Markdown-Überblick, ohne JSON oder
Trennzeichen.
""".strip(),
    "it": """
Documenti una cartella di lavoro Excel per un lettore aziendale.
Scrivi una sintesi concisa in Markdown con queste sezioni:
1. **Scopo** — il ruolo apparente della cartella, solo se il dossier lo
 conferma;
2. **Struttura** — i suoi fogli e la distribuzione dei calcoli;
3. **Flusso di calcolo** — i principali motivi di formule, nomi definiti e
 collegamenti;
4. **Automazione e limiti** — VBA, riferimenti esterni, avvisi e limiti
 dell'analisi;
5. **Questioni da validare** — al massimo cinque punti concreti non
 determinabili.
Usa solo i fatti presenti nel dossier deterministico. Non dedurre uno scopo
aziendale dai soli nomi dei fogli. Scrivi «non determinato dalla derivazione»
quando manca un'informazione.
Rispondi SOLO con la sintesi Markdown, senza JSON né delimitatori.
""".strip(),
    "pt": """
Documente uma pasta de trabalho do Excel para um leitor de negócio.
Redija uma síntese concisa em Markdown com estas secções:
1. **Finalidade** — o papel aparente da pasta, apenas se o dossiê o confirmar;
2. **Estrutura** — as suas folhas e a distribuição dos cálculos;
3. **Fluxo de cálculo** — os principais padrões de fórmulas, nomes definidos e
 ligações;
4. **Automação e limites** — VBA, referências externas, avisos e limites da
 análise;
5. **Questões a validar** — no máximo cinco pontos concretos indetermináveis.
Use apenas as informações do dossiê determinista. Não deduza uma finalidade de
negócio apenas a partir dos nomes das folhas. Escreva «não determinado pela
linhagem» quando faltar informação.
Responda APENAS com a síntese Markdown, sem JSON nem delimitadores.
""".strip(),
    "nl": """
Je documenteert een Excel-werkmap voor een zakelijke lezer.
Schrijf een beknopt Markdown-overzicht met deze secties:
1. **Doel** — de kennelijke rol van de werkmap, alleen als het dossier dit
 staaft;
2. **Structuur** — de bladen en de verdeling van de berekeningen;
3. **Berekeningsstroom** — belangrijke formulepatronen, gedefinieerde namen en
 koppelingen;
4. **Automatisering en beperkingen** — VBA, externe verwijzingen,
 waarschuwingen en analysegrenzen;
5. **Te valideren vragen** — maximaal vijf concrete, niet vast te stellen
 punten.
Gebruik uitsluitend de gegevens uit het deterministische dossier. Leid geen
zakelijk doel af uit alleen de bladnamen. Schrijf "niet bepaald door de
herkomst" wanneer informatie ontbreekt.
Antwoord UITSLUITEND met het Markdown-overzicht, zonder JSON of
scheidingstekens.
""".strip(),
    "ja": """
あなたはビジネス読者向けに Excel ブックを文書化します。
次のセクションからなる簡潔な Markdown の概要を作成してください。
1. **目的** — ドシエが裏付ける場合に限り、ブックの明らかな役割。
2. **構成** — シートと計算の配分。
3. **計算の流れ** — 主要な数式パターン、定義された名前、リンク。
4. **自動化と留意点** — VBA、外部参照、警告、分析の限界。
5. **確認すべき点** — 特定できない具体的な項目を最大 5 件。
決定論的なドシエにある事実のみを使用すること。シート名だけから業務上の
目的を推測しないこと。情報が不足している場合は「系統からは特定できません」
と書くこと。
Markdown の概要のみで回答し、JSON や区切り記号は使わないこと。
""".strip(),
    "zh": """
你为业务读者记录一个 Excel 工作簿。
撰写一份简明的 Markdown 概览，包含以下部分：
1. **目的** — 仅在档案支持时，说明该工作簿的明显作用；
2. **结构** — 其工作表以及计算的分布；
3. **计算流程** — 主要的公式模式、定义的名称和链接；
4. **自动化与局限** — VBA、外部引用、警告以及分析的限制；
5. **待确认的问题** — 最多五个无法确定的具体事项。
仅使用确定性档案中的事实。不要仅凭工作表名称推断业务目的。
信息缺失时请写“无法由血缘确定”。
仅以 Markdown 概览作答，不要使用 JSON 或分隔符。
""".strip(),
}


class AiDocError(RuntimeError):
    pass


# ──────────────────────────────────────────────
# Token accounting
# ──────────────────────────────────────────────

# Runs of CJK ideographs, kana and hangul, which tokenizers split at roughly one
# token per character while ``\w+`` would swallow a whole sentence as one word.
_CJK_RE = re.compile(r"[぀-ヿ㐀-䶿一-鿿豈-﫿가-힯]")
_WORD_RE = re.compile(r"\w+")


def estimate_tokens(text: str) -> int:
    """Approximate the token count of ``text``.

    Only a fallback: :class:`TokenUsage` prefers the counts the provider
    reports. Latin script is counted as words × 4/3 (the usual 1 token ≈ 0.75
    words ratio); CJK characters are counted individually, because a Japanese
    or Chinese sentence carries no spaces and would otherwise register as a
    single word.

    >>> estimate_tokens("the quick brown fox jumps")
    6
    >>> estimate_tokens("")
    0
    """
    if not text:
        return 0
    cjk = len(_CJK_RE.findall(text))
    latin_words = len(_WORD_RE.findall(_CJK_RE.sub(" ", text)))
    return cjk + latin_words * 4 // 3


@dataclass
class TokenUsage:
    """Tokens consumed by one or more documentation requests.

    ``estimated`` is ``True`` as soon as any request in the tally had to be
    approximated by :func:`estimate_tokens` instead of being reported by the
    provider — treat such a total as an order of magnitude, not a bill.
    """

    input_tokens: int = 0
    output_tokens: int = 0
    requests: int = 0
    estimated: bool = False
    model: str = ""
    provider: str = ""

    @property
    def total(self) -> int:
        return self.input_tokens + self.output_tokens

    def add(self, other: TokenUsage) -> None:
        """Accumulate ``other`` in place, keeping the model/provider labels."""
        self.input_tokens += other.input_tokens
        self.output_tokens += other.output_tokens
        self.requests += other.requests
        self.estimated = self.estimated or other.estimated
        self.model = self.model or other.model
        self.provider = self.provider or other.provider

    def __str__(self) -> str:
        about = "~" if self.estimated else ""
        where = f" [{self.provider}/{self.model}]" if self.provider else ""
        return (
            f"{about}{self.total:,} tokens "
            f"({about}{self.input_tokens:,} in + {about}{self.output_tokens:,} out) "
            f"over {self.requests} request(s){where}"
        )


# ──────────────────────────────────────────────
# Provider abstraction
# ──────────────────────────────────────────────


@runtime_checkable
class LLMProvider(Protocol):
    """Minimal protocol: system + user prompt → text response."""

    def generate(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> str: ...


@runtime_checkable
class UsageReportingProvider(Protocol):
    """A provider that also reports what the call consumed.

    Optional: the built-in Gemini and OpenAI-compatible clients implement it so
    that token counts come from the API rather than from an approximation.
    Custom providers only need :class:`LLMProvider`.
    """

    def generate_with_usage(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> tuple[str, TokenUsage]: ...


def _generate(
    llm: LLMProvider,
    system_prompt: str,
    user_prompt: str,
    max_tokens: int | None = None,
) -> tuple[str, TokenUsage]:
    """Call ``llm``, returning its text and what the call cost.

    Providers reporting real usage are preferred; anything else is estimated.
    """
    if isinstance(llm, UsageReportingProvider):
        return llm.generate_with_usage(
            system_prompt, user_prompt, temperature=0.2, max_tokens=max_tokens
        )
    text = llm.generate(
        system_prompt, user_prompt, temperature=0.2, max_tokens=max_tokens
    )
    return text, TokenUsage(
        input_tokens=estimate_tokens(system_prompt) + estimate_tokens(user_prompt),
        output_tokens=estimate_tokens(text or ""),
        requests=1,
        estimated=True,
    )


#: Anything accepted as ``provider=``: an :class:`LLMProvider` (any object with
#: a ``generate`` method) or a plain callable with the same signature.
ProviderLike = LLMProvider | Callable[..., str]


class _CallableProvider:
    """Adapt a plain callable to the :class:`LLMProvider` protocol."""

    def __init__(self, fn: Callable[..., str]):
        self._fn = fn

    def generate(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> str:
        return self._fn(system_prompt, user_prompt, temperature=temperature)


def _as_provider(provider: ProviderLike) -> LLMProvider:
    """Normalize a user-supplied provider into an :class:`LLMProvider`."""
    if isinstance(provider, LLMProvider):
        return provider
    if callable(provider):
        return _CallableProvider(provider)
    raise AiDocError(
        "provider must expose a generate(system_prompt, user_prompt, *, "
        f"temperature) method or be callable, got {type(provider).__name__}"
    )


def _resolve_provider(
    *,
    provider: ProviderLike | None = None,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
) -> LLMProvider:
    """Resolve which provider to use.

    Priority:
    1. Explicit `provider` (custom callable or LLMProvider instance)
    2. `base_url` set → OpenAI-compatible client
    3. Google API key available → Gemini client (backward-compatible default)
    """
    if provider is not None:
        return _as_provider(provider)

    # OpenAI-compatible endpoint (Ollama, vLLM, LM Studio, OpenAI, etc.)
    base = base_url or os.getenv("LINEXCEL_AI_BASE_URL") or os.getenv("OPENAI_BASE_URL")
    if base:
        resolved_model = (
            model
            or os.getenv("LINEXCEL_AI_MODEL")
            or os.getenv("OPENAI_MODEL")
            or "gpt-4o-mini"
        )
        return _OpenAICompatProvider(
            base_url=base, api_key=api_key, model=resolved_model
        )

    # Google Gemini (default, backward-compatible)
    return _GeminiProvider(
        api_key=api_key, model=model or os.getenv("GEMINI_MODEL", DEFAULT_MODEL)
    )


class _GeminiProvider:
    """Google Gemini via google-genai."""

    def __init__(self, *, api_key: str | None = None, model: str = DEFAULT_MODEL):
        try:
            from google import genai
        except ImportError as exc:  # pragma: no cover
            raise AiDocError(
                "google-genai is not installed "
                "(pip install 'linexcel[ai]' or pip install google-genai)"
            ) from exc
        self._genai = genai
        self._api_key = (
            api_key or os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
        )
        if not self._api_key:
            raise AiDocError(
                "No Gemini API key provided: pass api_key=... or set "
                "GOOGLE_API_KEY in the environment"
            )
        self._client = genai.Client(api_key=self._api_key)
        self._model = model

    def generate(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> str:
        return self.generate_with_usage(
            system_prompt, user_prompt, temperature=temperature, max_tokens=max_tokens
        )[0]

    def generate_with_usage(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> tuple[str, TokenUsage]:
        prompt = system_prompt + "\n\n" + user_prompt
        config: dict[str, Any] = {"temperature": temperature}
        if max_tokens is not None:
            config["max_output_tokens"] = max_tokens
        try:
            response = self._client.models.generate_content(
                model=self._model,
                contents=prompt,
                config=config,
            )
            text = (response.text or "").strip()
        except Exception as exc:
            raise AiDocError(f"Gemini API call failed: {exc}") from exc
        return text, _usage_from(
            getattr(response, "usage_metadata", None),
            ("prompt_token_count", "candidates_token_count"),
            prompt,
            text,
            model=self._model,
            provider="gemini",
        )


class _OpenAICompatProvider:
    """OpenAI-compatible API client (works with Ollama, vLLM, LM Studio, etc.)."""

    def __init__(
        self, *, base_url: str, api_key: str | None = None, model: str = "gpt-4o-mini"
    ):
        try:
            from openai import OpenAI
        except ImportError as exc:  # pragma: no cover
            raise AiDocError("openai is not installed (pip install openai)") from exc
        self._client = OpenAI(
            api_key=api_key or os.getenv("OPENAI_API_KEY") or "ollama",  # no key needed
            base_url=base_url,
        )
        self._model = model

    def generate(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> str:
        return self.generate_with_usage(
            system_prompt, user_prompt, temperature=temperature, max_tokens=max_tokens
        )[0]

    def generate_with_usage(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> tuple[str, TokenUsage]:
        kwargs: dict[str, Any] = {"temperature": temperature}
        if max_tokens is not None:
            kwargs["max_tokens"] = max_tokens
        try:
            response = self._client.chat.completions.create(
                model=self._model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                **kwargs,
            )
            text = (response.choices[0].message.content or "").strip()
        except Exception as exc:
            raise AiDocError(f"OpenAI-compatible API call failed: {exc}") from exc
        return text, _usage_from(
            getattr(response, "usage", None),
            ("prompt_tokens", "completion_tokens"),
            system_prompt + "\n\n" + user_prompt,
            text,
            model=self._model,
            provider="openai-compatible",
        )


def _usage_from(
    reported: Any,
    fields: tuple[str, str],
    prompt: str,
    text: str,
    *,
    model: str,
    provider: str,
) -> TokenUsage:
    """Build a :class:`TokenUsage` from what the API reported, else estimate.

    Local runtimes behind an OpenAI-compatible endpoint do not always populate
    the usage block, so each field falls back independently.
    """
    prompt_field, completion_field = fields
    reported_in = getattr(reported, prompt_field, None) if reported else None
    reported_out = getattr(reported, completion_field, None) if reported else None
    estimated = reported_in is None or reported_out is None
    return TokenUsage(
        input_tokens=estimate_tokens(prompt) if reported_in is None else reported_in,
        output_tokens=estimate_tokens(text) if reported_out is None else reported_out,
        requests=1,
        estimated=estimated,
        model=model,
        provider=provider,
    )


# ──────────────────────────────────────────────
# Dossier builders (unchanged)
# ──────────────────────────────────────────────


def build_dossier(graph: dict[str, Any], node_id: str) -> dict[str, Any] | None:
    """
    Deterministic dossier for a node: everything the AI is allowed to use.
    """
    nodes = {n["id"]: n for n in graph["nodes"]}
    node = nodes.get(node_id)
    if node is None:
        return None
    precedents, dependents = [], []
    for e in graph["edges"]:
        if e["target"] == node_id:
            src = nodes.get(e["source"], {})
            precedents.append(_neighbor(src, e))
        elif e["source"] == node_id:
            dst = nodes.get(e["target"], {})
            dependents.append(_neighbor(dst, e))
    dossier = {
        "node_id": node_id,
        "kind": node.get("kind"),
        "sheet": node.get("sheet"),
        "address": node.get("addr"),
        "formula": node.get("formula"),
        "r1c1_form": node.get("r1c1"),
        "group_cells": node.get("count"),
        "extent": node.get("bbox"),
        "computed_value": node.get("value"),
        "value_samples": node.get("samples"),
        "decomposition": _compact_steps(node.get("steps")),
        "precedents": precedents[:30],
        "dependents": dependents[:30],
    }
    if node.get("kind") == "vba":
        dossier["vba"] = {
            "module": node.get("module"),
            "procedure": node.get("proc"),
            "type": node.get("procKind"),
            "code": (node.get("code") or "")[:2500],
        }
    return dossier


def _neighbor(other: dict, edge: dict) -> dict:
    return {
        "id": other.get("id"),
        "kind": other.get("kind"),
        "label": other.get("label"),
        "edge_kind": edge.get("kind"),
        "value": other.get("value"),
        "formula": other.get("formula"),
    }


def _compact_steps(step: dict | None) -> dict | None:
    if step is None:
        return None
    out = {
        "expression": step.get("expr"),
        "operation": step.get("label"),
        "value": step.get("value") if step.get("evaluated") else "not evaluated",
    }
    if step.get("inputs"):
        out["inputs"] = step["inputs"]
    children = [_compact_steps(c) for c in step.get("children", [])]
    if children:
        out["sub_steps"] = children
    return out


def build_workbook_dossier(graph: dict[str, Any]) -> dict[str, Any]:
    """Return a compact, deterministic dossier for a whole-workbook overview."""
    nodes = graph.get("nodes", [])
    meta = graph.get("meta", {})
    stats = meta.get("stats", {})
    sheet_stats = stats.get("sheets", [])
    nodes_by_sheet: dict[str, dict[str, int]] = {}
    for node in nodes:
        sheet = node.get("sheet")
        if not sheet:
            continue
        kinds = nodes_by_sheet.setdefault(sheet, {})
        kind = node.get("kind", "unknown")
        kinds[kind] = kinds.get(kind, 0) + 1

    sheets = [
        {
            "name": sheet.get("name"),
            "dimensions": {"rows": sheet.get("rows"), "columns": sheet.get("cols")},
            "formula_cells": sheet.get("formulaCells", 0),
            "lineage_nodes": nodes_by_sheet.get(sheet.get("name"), {}),
        }
        for sheet in sheet_stats
    ]
    formula_patterns = sorted(
        (
            {
                "sheet": node.get("sheet"),
                "address": node.get("addr"),
                "formula": node.get("formula"),
                "cells": node.get("count", 1),
                "extent": node.get("bbox"),
            }
            for node in nodes
            if node.get("kind") in {"cell", "group"}
        ),
        key=lambda item: item["cells"],
        reverse=True,
    )[:20]
    defined_names = [
        {"name": node.get("label"), "targets": node.get("targets", [])}
        for node in nodes
        if node.get("kind") == "name"
    ]
    vba = [
        {
            "module": node.get("module"),
            "procedure": node.get("proc"),
            "type": node.get("procKind"),
        }
        for node in nodes
        if node.get("kind") == "vba"
    ]
    opaque_references = [
        node.get("label") for node in nodes if node.get("kind") == "opaque"
    ]
    return {
        "filename": meta.get("filename"),
        "analysis": {
            "formula_cells": stats.get("totalFormulas", 0),
            "lineage_nodes": stats.get("totalNodes", 0),
            "lineage_edges": stats.get("totalEdges", 0),
            "grouped_patterns": stats.get("groupedPatterns", 0),
        },
        "sheets": sheets,
        "formula_patterns": formula_patterns,
        "defined_names": defined_names,
        "vba_procedures": vba,
        "external_or_unresolved_references": opaque_references,
        "warnings": meta.get("warnings", []),
    }


# ──────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────


def document_workbook(
    graph: dict[str, Any],
    *,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
    provider: ProviderLike | None = None,
    language: str = "en",
    usage: TokenUsage | None = None,
    max_tokens: int | None = None,
) -> str:
    """Generate a Markdown overview grounded in the workbook dossier.

    Provider resolution (first match wins):
    1. `provider` — custom LLMProvider instance or callable
    2. `base_url` or `LINEXCEL_AI_BASE_URL` — OpenAI-compatible endpoint
    3. Google Gemini (default, backward-compatible)

    If a :class:`TokenUsage` is passed as ``usage``, what the call consumed is
    accumulated into it.
    """
    if language not in _LANGUAGES:
        raise ValueError(f"Unsupported language: {language!r}. Use one of {_LANGUAGES}")
    dossier = build_workbook_dossier(graph)
    blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) > MAX_WORKBOOK_DOSSIER_CHARS:
        dossier["formula_patterns"] = dossier["formula_patterns"][:5]
        dossier["vba_procedures"] = dossier["vba_procedures"][:10]
        blob = json.dumps(dossier, ensure_ascii=False, default=str)
    llm = _resolve_provider(
        provider=provider, model=model, api_key=api_key, base_url=base_url
    )
    system = _WORKBOOK_SYSTEM[language]
    user = "Workbook dossier (deterministic, extracted from workbook):\n" + blob
    try:
        text, call_usage = _generate(llm, system, user, max_tokens=max_tokens)
    except AiDocError:
        raise
    except Exception as exc:
        raise AiDocError(f"AI documentation failed: {exc}") from exc
    if usage is not None:
        usage.add(call_usage)
    return text


def document_nodes(
    graph: dict[str, Any],
    node_ids: list[str],
    *,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
    provider: ProviderLike | None = None,
    language: str = "en",
    max_workers: int = 4,
    usage: TokenUsage | None = None,
    max_tokens: int | None = None,
) -> dict[str, str]:
    """Document the requested nodes, returns {node_id: markdown}.

    Provider resolution is the same as :func:`document_workbook`.

    Nodes are documented concurrently (``max_workers`` in-flight requests;
    raise it if the provider's rate limits allow). Documenting a large
    workbook is a long, often billed operation, so a node that fails does not
    discard the ones that succeeded: the successful cards are returned and a
    :class:`UserWarning` reports how many nodes were dropped.
    :class:`AiDocError` is raised only when *every* node failed.

    If a :class:`TokenUsage` is passed as ``usage``, every successful call is
    accumulated into it — including those of a run that later fails, since
    tokens already spent are still billed.
    """
    if language not in _LANGUAGES:
        raise ValueError(f"Unsupported language: {language!r}. Use one of {_LANGUAGES}")
    if max_workers < 1:
        raise ValueError("max_workers must be >= 1")
    llm = _resolve_provider(
        provider=provider, model=model, api_key=api_key, base_url=base_url
    )
    system = _SYSTEM[language]
    docs: dict[str, str] = {}
    dossiers = []
    for nid in node_ids:
        d = build_dossier(graph, nid)
        if d is not None:
            blob = json.dumps(d, ensure_ascii=False, default=str)
            if len(blob) > MAX_DOSSIER_CHARS:
                d["decomposition"] = "truncated (very long formula)"
                blob = json.dumps(d, ensure_ascii=False, default=str)
            dossiers.append((nid, blob))
    if not dossiers:
        return docs

    def _doc_one(nid_blob: tuple[str, str]) -> tuple[str, str, TokenUsage]:
        nid, blob = nid_blob
        user = "Lineage dossier (deterministic, extracted from workbook):\n" + blob
        text, call_usage = _generate(llm, system, user, max_tokens=max_tokens)
        return nid, text or "(AI returned empty response)", call_usage

    failures: list[tuple[str, Exception]] = []
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(_doc_one, d): d[0] for d in dossiers}
        # Accumulating here rather than inside the workers keeps `usage` free of
        # races: this loop is the single consumer of the pool's results.
        for fut in as_completed(futures):
            node_id = futures[fut]
            try:
                nid, text, call_usage = fut.result()
            except Exception as exc:
                failures.append((node_id, exc))
                continue
            docs[nid] = text
            if usage is not None:
                usage.add(call_usage)

    if failures and not docs:
        node_id, exc = failures[0]
        raise AiDocError(
            f"AI documentation failed for all {len(failures)} nodes "
            f"(first error, node {node_id}): {exc}"
        ) from exc
    if failures:
        node_id, exc = failures[0]
        warnings.warn(
            f"AI documentation failed for {len(failures)} of {len(dossiers)} "
            f"nodes; {len(docs)} cards returned. First error (node {node_id}): "
            f"{exc}",
            UserWarning,
            stacklevel=2,
        )
    return docs
