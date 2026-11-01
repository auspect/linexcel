"""Source-only expansion of supported 3D aggregate arguments."""

from linexcel.refs import parse_ref, split_sheet_prefix

# Excel's documented reference-taking aggregate family. Expansion preserves
# reference arguments (including text/error semantics), not precomputed sums.
# https://support.microsoft.com/en-us/excel/get-started/overview-of-formulas-in-excel
AGGREGATES_3D = {
    "SUM",
    "AVERAGE",
    "AVERAGEA",
    "COUNT",
    "COUNTA",
    "MIN",
    "MINA",
    "MAX",
    "MAXA",
    "PRODUCT",
    "STDEV",
    "STDEVP",
    "STDEV.S",
    "STDEV.P",
    "STDEVA",
    "STDEVPA",
    "VAR",
    "VARP",
    "VAR.S",
    "VAR.P",
    "VARA",
    "VARPA",
}


def sheet_span(text, sheets):
    """Return source-order rectangles; None means this is not a local 3D token."""
    prefix, body = split_sheet_prefix(text)
    if prefix is None or "[" in prefix or ":" not in prefix:
        return None
    endpoints = prefix.split(":")
    rect = parse_ref(body)
    if len(endpoints) != 2 or rect is None:
        raise ValueError("Invalid 3D sheet span: " + text)
    positions = {sheet.casefold(): i for i, sheet in enumerate(sheets)}
    if any(endpoint.casefold() not in positions for endpoint in endpoints):
        raise ValueError("3D sheet endpoint is unavailable: " + text)
    start, end = sorted(positions[endpoint.casefold()] for endpoint in endpoints)
    from dataclasses import replace

    return [replace(rect, sheet=sheet) for sheet in sheets[start : end + 1]]


def aggregate_arguments(tokens):
    """Tokens which are complete arguments of supported 3D aggregate calls."""
    stack = []
    eligible = set()
    meaningful = [
        (i, token) for i, token in enumerate(tokens) if token.type != "WHITE-SPACE"
    ]
    for position, (index, token) in enumerate(meaningful):
        if token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "OPEN":
            stack.append(
                token.value[:-1].upper().removeprefix("_XLFN.")
                if token.type == "FUNC"
                else None
            )
            continue
        if token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "CLOSE":
            if stack:
                stack.pop()
            continue
        if token.type != "OPERAND" or token.subtype != "RANGE" or not stack:
            continue
        before = meaningful[position - 1][1] if position else None
        after = meaningful[position + 1][1] if position + 1 < len(meaningful) else None
        left = before is not None and (
            before.type == "FUNC"
            and before.subtype == "OPEN"
            or before.type == "SEP"
            and before.subtype == "ARG"
        )
        right = after is not None and (
            after.type == "FUNC"
            and after.subtype == "CLOSE"
            or after.type == "SEP"
            and after.subtype == "ARG"
        )
        if stack[-1] in AGGREGATES_3D and left and right:
            eligible.add(index)
    return eligible
