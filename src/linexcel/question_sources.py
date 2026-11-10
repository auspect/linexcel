"""Freeze named workbook sources before any model request or question token."""

from __future__ import annotations

import re
import unicodedata


class SourceSelectionError(ValueError):
    """The user's source selection cannot be resolved without guessing."""


def normalized_name(value: str) -> str:
    return unicodedata.normalize("NFC", value).casefold()


def select_question_sources(revisions, question, *, project=None, revision_id=None):
    """Intersect exact filename mentions with the caller's authorized selection.

    This recognizes literal workbook names, not natural-language intention.
    Missing or ambiguous names require a structured project/revision selection.
    Once that selection exists, a name absent from the library is only text, such
    as an external link target: it can neither widen nor replace the selection.
    No source path, owner, content or model participates in this resolution.
    """
    explicit = bool(project or revision_id)
    allowed = [
        item
        for item in revisions
        if (not project or item["projectId"] == project)
        and (not revision_id or item["id"] == revision_id)
    ]
    text = normalized_name(question)
    by_name = {}
    for item in revisions:
        by_name.setdefault(normalized_name(item["filename"]), []).append(item)
    spans, mentioned = [], []
    extension = r"\.(?:xlsx|xlsm|xlsb|xlam|xla|xls)"
    quoted_names = []
    for opening, closing in [
        ('"', '"'),
        ("“", "”"),
        ("«", "»"),
        ("'", "'"),
        ("‘", "’"),
        ("[", "]"),
        ("`", "`"),
    ]:
        escaped_close = re.escape(closing)
        if closing in {"'", "’"}:
            content = (
                r"(?:[^"
                + escaped_close
                + r"\r\n]|(?<=\w)"
                + escaped_close
                + r"(?=\w))+?"
            )
        else:
            content = r"[^" + escaped_close + r"\r\n]+?"
        pattern = (
            r"(?<!\w)"
            + re.escape(opening)
            + "("
            + content
            + extension
            + r")[ \t]*"
            + re.escape(closing)
            + ("" if closing == "]" else r"(?!\w)")
        )
        quoted_names.extend(re.finditer(pattern, text))
    # Resolve a complete quoted name before looking for known substrings. In
    # particular, "Missing Budget.xlsx" must never select the known Budget.xlsx.
    for match in sorted(
        quoted_names, key=lambda item: item.end() - item.start(), reverse=True
    ):
        if any(start <= match.start() and match.end() <= end for start, end in spans):
            continue
        name = match.group(1).strip()
        if name not in by_name and explicit:
            # Consume the span so a known suffix inside it is not selected.
            spans.append(match.span())
            continue
        if name not in by_name:
            raise SourceSelectionError(
                "Un classeur cité dans la question n’est pas disponible dans cette "
                "bibliothèque. Importez-le ou choisissez explicitement la source."
            )
        spans.append(match.span())
        if name not in mentioned:
            mentioned.append(name)
    for name in sorted(by_name, key=len, reverse=True):
        pattern = r"(?<![\w.-])" + re.escape(name) + r"(?![\w.-])"
        for match in re.finditer(pattern, text):
            if any(
                start <= match.start() and match.end() <= end for start, end in spans
            ):
                continue
            spans.append(match.span())
            if name not in mentioned:
                mentioned.append(name)

    # Remove recognized names before checking quoted full names and bare tokens.
    # This avoids treating 'FY26.xlsx' inside 'Budget FY26.xlsx' as another source.
    remainder = list(text)
    for start, end in spans:
        remainder[start:end] = " " * (end - start)
    remainder = "".join(remainder)
    bare = r"(?<![\w.-])[\w.-]+" + extension + r"(?![\w.-])"
    if not explicit and re.search(bare, remainder):
        raise SourceSelectionError(
            "Un classeur cité dans la question n’est pas disponible dans cette "
            "bibliothèque. Importez-le ou choisissez explicitement la source."
        )

    if mentioned:
        selected = []
        for name in mentioned:
            candidates = [
                item for item in allowed if normalized_name(item["filename"]) == name
            ]
            if not candidates:
                raise SourceSelectionError(
                    "La source citée ne correspond pas au classeur ou à la révision "
                    "sélectionnée. Choisissez la source voulue avant de rechercher."
                )
            if len(candidates) > 1:
                raise SourceSelectionError(
                    "Plusieurs classeurs portent le nom cité. Choisissez un classeur "
                    "ou une révision dans les options de recherche."
                )
            selected.extend(candidates)
        allowed = selected
    if not allowed:
        raise SourceSelectionError("Aucune révision importée dans cette portée")
    return allowed, {
        "kind": "explicit"
        if project or revision_id
        else "question_filename"
        if mentioned
        else "library",
        "filenames": list(dict.fromkeys(item["filename"] for item in allowed)),
    }
