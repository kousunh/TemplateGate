from __future__ import annotations

from ..core.messages import CATALOG, ENGLISH, Translator, say
from ..core.model import CheckResult


def human(value, limit: int = 80, t: Translator = ENGLISH) -> str:
    """A value as a person reads it, not as Python writes it.

    The report is read by whoever approves the change, and `None`, quoted
    strings and dict braces are noise to them — worse, they make a real value
    like the string "None" impossible to tell from an absent one.
    """
    if value is None:
        return t("value.none")
    if isinstance(value, str):
        text = value if value else t("value.empty")
    elif isinstance(value, bool):
        text = t("value.yes") if value else t("value.no")
    elif isinstance(value, dict):
        text = ", ".join(f"{name} {human(item, limit, t)}"
                         for name, item in value.items())
    elif isinstance(value, (list, tuple, set)):
        text = ", ".join(human(item, limit, t) for item in value)
    else:
        text = str(value)
    return text if len(text) <= limit else text[:limit] + "..."


def _worth_showing(change) -> bool:
    """Whether an old -> new line tells the reader anything.

    It does not when a named delta already spells the change out in the
    detail, when both sides are absent, or when the two sides render the same
    — "present -> present" fills two columns and says nothing, and the detail
    sentence beside it is doing all the work.
    """
    if change.old is None and change.new is None:
        return False
    if human(change.old) == human(change.new):
        return False
    # A named delta is only redundant if the detail is actually there to say
    # it; suppressing on shape alone would leave changes with nothing at all
    # to explain them.
    if isinstance(change.old, dict) or isinstance(change.new, dict):
        return not change.detail
    return True


def _grouped(violations):
    """Consecutive violations that describe one edit, folded together.

    Only the human-facing reports collapse; the JSON keeps every change, and
    the policy has already judged each of them separately.
    """
    runs: list[tuple[str, list]] = []
    for violation in violations:
        group = violation.change.group
        if group and runs and runs[-1][0] == group:
            runs[-1][1].append(violation)
        else:
            runs.append((group, [violation]))
    return runs


def verdict_of(violation, t: Translator) -> str:
    """Why this change is a violation, in the reader's language.

    Derived from the rule rather than translated from the stored message:
    that one is English by construction and belongs to the JSON contract, so
    it must not move.
    """
    key = f"verdict.{violation.rule}"
    if key in CATALOG["en"]:
        return t(key, attribute=violation.change.attribute)
    return violation.message


def _span(members) -> str:
    first, last = members[0].change, members[-1].change
    if first.location == last.location:
        return first.location
    return f"{first.location}..{last.location}"


def render_text(result: CheckResult, lang: str = "en") -> str:
    t = Translator(lang)
    status = t("report.pass") if result.passed else t("report.fail")
    # Pad the two labels to a common width rather than hardcoding the space
    # English needs: "baseline" is a character shorter than "candidate", but
    # 変更前 and 変更後 are the same length and a stray space just looks wrong.
    before, after = t("report.baseline"), t("report.candidate")
    width = max(len(before), len(after))
    lines = [
        t("report.title", status=status),
        f"  {before:<{width}}: {result.baseline}",
        f"  {after:<{width}}: {result.candidate}",
        "  " + t("report.summary", total=len(result.changes),
                 allowed=len(result.allowed),
                 violations=len(result.violations)),
    ]
    for role, reason in (result.meta.get("degraded") or {}).items():
        lines.append("  ! " + t("damaged.line", role=t(f"role.{role}"),
                                reason=reason))
        lines.append("    " + t("damaged.parts_only"))
    if result.warnings:
        lines.append("")
        lines.append(t("report.warnings"))
        for warning in result.warnings:
            lines.append(f"  ! {say(warning, t)}")
    if result.violations:
        lines.append("")
        if result.meta.get("policy_mode") == "review_only":
            lines.append(t("report.violations.review_only"))
        else:
            lines.append(t("report.violations"))
        for group, members in _grouped(result.violations):
            if group:
                lines.append(
                    f"  [{t('severity.' + members[0].severity)}] {_span(members)}: "
                    + t("report.cascade", reason=say(members[0].change.group, t),
                     count=len(members)))
                continue
            v = members[0]
            lines.append(f"  [{t('severity.' + v.severity)}] {v.change.location} "
                         f"({v.change.attribute}): {verdict_of(v, t)}")
            if _worth_showing(v.change):
                lines.append("      " + t("report.change_line",
                                          old=human(v.change.old, 80, t),
                                          new=human(v.change.new, 80, t)))
            detail = say(v.change.detail, t)
            if detail:
                lines.append(f"      {detail}")
    if result.semantic_mode != "off":
        lines.append("")
        lines.append(t("report.semantic", mode=result.semantic_mode))
        for f in result.semantic_findings:
            lines.append(f"  [{f.verdict}] {f.check}: {f.message}")
    return "\n".join(lines)
