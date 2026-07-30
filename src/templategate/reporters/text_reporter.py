from __future__ import annotations

from ..core.model import CheckResult


def human(value, limit: int = 80) -> str:
    """A value as a person reads it, not as Python writes it.

    The report is read by whoever approves the change, and `None`, quoted
    strings and dict braces are noise to them — worse, they make a real value
    like the string "None" impossible to tell from an absent one.
    """
    if value is None:
        return "none"
    if isinstance(value, str):
        text = value if value else "(empty)"
    elif isinstance(value, bool):
        text = "yes" if value else "no"
    elif isinstance(value, dict):
        text = ", ".join(f"{name} {human(item, limit)}"
                         for name, item in value.items())
    elif isinstance(value, (list, tuple, set)):
        text = ", ".join(human(item, limit) for item in value)
    else:
        text = str(value)
    return text if len(text) <= limit else text[:limit] + "..."


def _fmt(value) -> str:
    return human(value)


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


def render_text(result: CheckResult) -> str:
    lines = [
        f"TemplateGate: {'PASS' if result.passed else 'FAIL'}",
        f"  baseline : {result.baseline}",
        f"  candidate: {result.candidate}",
        f"  changes: {len(result.changes)} total, "
        f"{len(result.allowed)} allowed, {len(result.violations)} violations",
    ]
    for role, reason in (result.meta.get("degraded") or {}).items():
        lines.append(f"  ! the {role} document is damaged: {reason}")
        lines.append("    only its package parts could be compared")
    if result.warnings:
        lines.append("")
        lines.append("Policy warnings:")
        for warning in result.warnings:
            lines.append(f"  ! {warning}")
    if result.violations:
        lines.append("")
        if result.meta.get("policy_mode") == "review_only":
            lines.append("Violations (mode: review_only — reported, not blocking):")
        else:
            lines.append("Violations:")
        for group, members in _grouped(result.violations):
            if group:
                first, last = members[0].change, members[-1].change
                span = (first.location if first.location == last.location
                        else f"{first.location}..{last.location}")
                lines.append(
                    f"  [{members[0].severity}] {span}: content shifted because "
                    f"{group} — {len(members)} knock-on changes collapsed; "
                    "the JSON report lists each one")
                continue
            v = members[0]
            lines.append(f"  [{v.severity}] {v.change.location} ({v.change.attribute}): {v.message}")
            # A named delta already reads "numfmt #,##0 -> #,##0," in the
            # detail; repeating it as old=/new= says the same thing twice.
            spelled_out = isinstance(v.change.old, dict) or isinstance(v.change.new, dict)
            if not spelled_out and (v.change.old is not None or v.change.new is not None):
                lines.append(f"      {_fmt(v.change.old)} -> {_fmt(v.change.new)}")
            if v.change.detail:
                lines.append(f"      {v.change.detail}")
    if result.semantic_mode != "off":
        lines.append("")
        lines.append(f"Semantic checks ({result.semantic_mode}):")
        for f in result.semantic_findings:
            lines.append(f"  [{f.verdict}] {f.check}: {f.message}")
    return "\n".join(lines)
