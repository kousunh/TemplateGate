"""Warn when a policy rule names something the baseline does not have.

A selector that matches nothing is silent in both directions: an allow rule
that reaches nothing simply never allows, and a *protect* rule that reaches
nothing protects nothing while looking like it protects something.  A typo in
a sheet name is indistinguishable from a policy that is deliberately narrow.

The check is deliberately timid.  It only speaks up about containers it can
enumerate from the baseline — a sheet name, a paragraph count — and says
nothing about ranges inside them, because a policy that allows edits in a
currently-empty range is perfectly reasonable and a warning there would be
wrong.  A wrong warning costs more than a missing one: it teaches people to
ignore the warnings.
"""

from __future__ import annotations

import re

from .messages import message
from .policy import Policy
from .selector import normalize

# Selectors that address change namespaces rather than document content.
# Nothing about the baseline says whether they will match, so they are never
# questioned.
_NAMESPACE_PREFIXES = ("sheet:", "name:", "package#", "workbook#")
_NAMESPACE_EXACT = frozenset({"*", "vba", "body"})

_PARAGRAPH = re.compile(r"p(\d+)(?:-(\d+))?$")
_INDEXED = re.compile(r"(table|sdt|textbox|section)(\d+)")


def _listed(names):
    shown = sorted(names)
    if not shown:
        return message("liveness.listed_none")
    if len(shown) > 6:
        return message("liveness.listed_more",
                       names=", ".join(repr(n) for n in shown[:6]))
    return message("liveness.listed", names=", ".join(repr(n) for n in shown))


def _excel_complaint(selector: str, snapshot: dict) -> str | None:
    sheets = set(snapshot.get("sheets", {}))
    if not sheets:
        return None
    # Reuse the selector's own parser so quoting and separators agree.
    from .selector import _split_sheet

    sheet, separator, _rest = _split_sheet(selector)
    if not sheet:
        return None
    if normalize(sheet) in {normalize(name) for name in sheets}:
        return None
    if separator == "" and selector != sheet:
        return None
    return message("liveness.no_sheet", sheet=sheet,
                   listed=_listed(sheets))


def _word_complaint(selector: str, snapshot: dict) -> str | None:
    paragraphs = len(snapshot.get("paragraphs", []))
    match = _PARAGRAPH.fullmatch(selector)
    if match:
        low = int(match.group(1))
        if low > paragraphs:
            return message("liveness.no_paragraph.one" if paragraphs == 1
                           else "liveness.no_paragraph.many",
                           available=paragraphs, wanted=low)
        return None

    match = _INDEXED.fullmatch(selector)
    if match:
        kind, index = match.group(1), int(match.group(2))
        available = {
            "table": len(snapshot.get("tables", [])),
            "section": len(snapshot.get("sections", [])),
            "sdt": sum(1 for key in snapshot.get("blocks", {})
                       if key.startswith("sdt")),
            "textbox": sum(1 for key in snapshot.get("blocks", {})
                           if key.startswith("textbox")),
        }[kind]
        if index > available:
            return message("liveness.no_indexed.one" if available == 1
                           else "liveness.no_indexed.many",
                           available=available, kind=kind, wanted=index)
    return None


def _complaint(selector: str, snapshot: dict) -> str | None:
    selector = selector.strip()
    if selector in _NAMESPACE_EXACT or selector.startswith(_NAMESPACE_PREFIXES):
        return None
    if snapshot.get("degraded"):
        return None  # a document we could not fully read proves nothing
    if snapshot.get("target") == "excel":
        return _excel_complaint(selector, snapshot)
    return _word_complaint(selector, snapshot)


def policy_warnings(policy: Policy, baseline: dict) -> list[str]:
    """Rules that cannot match anything in the baseline, in plain words."""
    messages: list[str] = []
    for kind, rules in (("allow", policy.allow), ("protect", policy.protect)):
        for number, rule in enumerate(rules, start=1):
            complaint = _complaint(rule.selector, baseline)
            if complaint:
                messages.append(message("liveness.dead_rule", kind=kind,
                                        number=number, selector=rule.selector,
                                        complaint=complaint))
    return messages
