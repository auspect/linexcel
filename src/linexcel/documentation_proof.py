"""Insert exact formula evidence without asking a model to transcribe it."""

from __future__ import annotations

import hashlib
import json
import re
from typing import Any

_TOKEN_BODY = (
    r"(?:SOURCE_FORMULA|STORED_XML_FORMULA|CANONICAL_FORMULA)(?::[^{}\r\n]{1,256})?"
)
_TOKENS = re.compile(
    r"\{\{(SOURCE_FORMULA|STORED_XML_FORMULA|CANONICAL_FORMULA)(?::([^{}\r\n]{1,256}))?\}\}"
)
_INLINE_TOKEN = re.compile(
    r"(?<!`)(?P<ticks>`+)(?P<token>\{\{" + _TOKEN_BODY + r"\}\})(?P=ticks)(?!`)"
)
_TOKEN_FENCE = re.compile(
    r"(?m)^[ \t]*(?P<ticks>`{3,})[^\r\n]*\r?\n"
    r"(?P<body>[\s\S]*?)^[ \t]*(?P=ticks)[ \t]*$"
)


def render_formula_tokens(
    markdown: str, dossier: dict[str, Any]
) -> tuple[str, list[dict[str, Any]]]:
    """Insert supplied focal fields or one exact retained overview example.

    This is deterministic insertion, not a check or repair of generated prose.
    Reports retain the original response and a separate insertion ledger. No
    formula is inferred from a cache, decomposition or an unselected node.
    """
    nested = dossier.get("node")
    node = nested if isinstance(nested, dict) else dossier
    prefix = "/node" if isinstance(nested, dict) else ""
    camel = isinstance(nested, dict)
    provenance_key = "formulaProvenance" if camel else "formula_provenance"
    provenance = node.get(provenance_key)
    provenance = provenance if isinstance(provenance, dict) else {}
    fields = {
        "SOURCE_FORMULA": (
            "sourceFormula" if camel else "source_formula",
            "presented_original_or_derived_shared_translation",
        ),
        "STORED_XML_FORMULA": (
            provenance_key + "/xmlText",
            "decoded_stored_xml_without_added_equals",
        ),
        "CANONICAL_FORMULA": (
            "canonicalFormula" if camel else "canonical_formula",
            "canonical_expression",
        ),
    }
    insertions: list[dict[str, Any]] = []

    def unwrap_token_fence(match: re.Match[str]) -> str:
        body = _INLINE_TOKEN.sub(lambda m: m["token"], match["body"])
        if _TOKENS.search(body) and not _TOKENS.sub("", body).strip():
            return body
        return match[0]

    # Models sometimes wrap placeholders in code despite the prompt. Consume
    # token-only wrappers before inserting blocks; retain all other model text.
    markdown = _TOKEN_FENCE.sub(unwrap_token_fence, markdown)
    markdown = _INLINE_TOKEN.sub(lambda m: m["token"], markdown)

    def replace(match: re.Match[str]) -> str:
        token = match[1]
        field, role = fields[token]
        selected, selected_provenance, selected_prefix = node, provenance, prefix
        requested = match[2]
        if requested is not None:
            field = {
                "SOURCE_FORMULA": "source_formula",
                "STORED_XML_FORMULA": "formula_provenance/xmlText",
                "CANONICAL_FORMULA": "canonical_formula",
            }[token]
            patterns = dossier.get("formula_patterns")
            patterns = patterns if isinstance(patterns, list) else []
            candidates = [
                (index, item)
                for index, item in enumerate(patterns)
                if isinstance(item, dict) and item.get("node_id") == requested
            ]
            if len(candidates) == 1:
                index, selected = candidates[0]
                selected_prefix = f"/formula_patterns/{index}"
                selected_provenance = selected.get("formula_provenance") or {}
                if not isinstance(selected_provenance, dict):
                    selected_provenance = {}
            else:
                # Never borrow a neighboring example or choose an ambiguous ID.
                selected, selected_provenance = {}, {}
                selected_prefix = "/formula_patterns"
        value = (
            selected_provenance.get("xmlText")
            if token == "STORED_XML_FORMULA"
            else selected.get(field)
        )
        record: dict[str, Any] = {
            "token": match[0],
            "pointer": selected_prefix + "/" + field,
            "role": role,
            "status": "inserted"
            if isinstance(value, str) and value
            else "not_supplied",
        }
        if requested is not None:
            record["target_node_id"] = requested
        insertions.append(record)
        if record["status"] == "not_supplied":
            return "`" + field + "`: `not_supplied_in_sent_dossier`"
        assert isinstance(value, str)
        # A source containing Markdown delimiters or controls is displayed as
        # a JSON string, explicitly labelled; decoding it recovers exact text.
        encoded = "`" in value or any(ord(ch) < 32 and ch not in "\n\t" for ch in value)
        displayed = json.dumps(value, ensure_ascii=False) if encoded else value
        record.update(
            value=value,
            characters=len(value),
            utf8_bytes=len(value.encode("utf-8")),
            value_sha256=hashlib.sha256(value.encode("utf-8")).hexdigest(),
            representation="json_encoded_string" if encoded else "decoded_text",
        )
        label = "`" + field + "` · `" + role + "`"
        if encoded:
            label += " · `json_encoded_string`"
        return "\n\n" + label + "\n\n```text\n" + displayed + "\n```\n\n"

    return _TOKENS.sub(replace, markdown), insertions
