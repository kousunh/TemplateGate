from __future__ import annotations

from ..core.model import CheckResult
from .text_reporter import _grouped, _worth_showing, human


def _escape(text: str) -> str:
    """Keep a value from breaking out of its Markdown table cell."""
    return str(text).replace("|", "\\|").replace("\n", " ")


def _cell(value) -> str:
    return _escape(human(value, limit=60))


def render_markdown(result: CheckResult) -> str:
    status = "✅ PASS" if result.passed else "❌ FAIL"
    lines = [
        f"## TemplateGate: {status}",
        "",
        f"- Baseline: `{result.baseline}`",
        f"- Candidate: `{result.candidate}`",
        f"- Changes: {len(result.changes)} total / "
        f"{len(result.allowed)} allowed / {len(result.violations)} violations",
    ]
    for role, reason in (result.meta.get("degraded") or {}).items():
        lines += ["", f"> **The {role} document is damaged:** {_escape(reason)}  ",
                  "> Only its package parts could be compared."]
    if result.warnings:
        lines += ["", "### Policy warnings", ""]
        lines += [f"- {_escape(warning)}" for warning in result.warnings]
    if result.violations:
        lines += [""]
        if result.meta.get("policy_mode") == "review_only":
            lines += ["> `mode: review_only` — reported for review, not blocking.", ""]
        lines += [
            "| Severity | Location | Attribute | Old | New | Rule |",
            "|---|---|---|---|---|---|",
        ]
        for group, members in _grouped(result.violations):
            if group:
                first, last = members[0].change, members[-1].change
                span = (first.location if first.location == last.location
                        else f"{first.location}..{last.location}")
                lines.append(
                    f"| {members[0].severity} | `{_escape(span)}` "
                    f"| {_escape(first.attribute)} "
                    f"| {_escape(group)} | {len(members)} knock-on changes "
                    f"| {_escape(members[0].rule)} |"
                )
                continue
            v = members[0]
            if _worth_showing(v.change):
                before, after = _cell(v.change.old), _cell(v.change.new)
            else:
                # There is no Old/New worth printing, so the row carries the
                # explanation instead of two columns of "none".
                before, after = _escape(v.change.detail or v.message), ""
            lines.append(
                f"| {v.severity} | `{_escape(v.change.location)}` "
                f"| {_escape(v.change.attribute)} "
                f"| {before} | {after} "
                f"| {_escape(v.rule)} |"
            )
    if result.semantic_mode != "off" and result.semantic_findings:
        lines += ["", f"### Semantic checks ({result.semantic_mode})", ""]
        for f in result.semantic_findings:
            lines.append(f"- **{f.verdict}** — {f.check}: {f.message}")
    return "\n".join(lines)
