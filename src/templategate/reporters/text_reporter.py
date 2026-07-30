from __future__ import annotations

from ..core.model import CheckResult


def _fmt(value) -> str:
    text = repr(value)
    return text if len(text) <= 80 else text[:80] + "..."


def render_text(result: CheckResult) -> str:
    lines = [
        f"TemplateGate: {'PASS' if result.passed else 'FAIL'}",
        f"  baseline : {result.baseline}",
        f"  candidate: {result.candidate}",
        f"  changes: {len(result.changes)} total, "
        f"{len(result.allowed)} allowed, {len(result.violations)} violations",
    ]
    if result.violations:
        lines.append("")
        if result.meta.get("policy_mode") == "review_only":
            lines.append("Violations (mode: review_only — reported, not blocking):")
        else:
            lines.append("Violations:")
        for v in result.violations:
            lines.append(f"  [{v.severity}] {v.change.location} ({v.change.attribute}): {v.message}")
            if v.change.old is not None or v.change.new is not None:
                lines.append(f"      old={_fmt(v.change.old)} new={_fmt(v.change.new)}")
            if v.change.detail:
                lines.append(f"      {v.change.detail}")
    if result.semantic_mode != "off":
        lines.append("")
        lines.append(f"Semantic checks ({result.semantic_mode}):")
        for f in result.semantic_findings:
            lines.append(f"  [{f.verdict}] {f.check}: {f.message}")
    return "\n".join(lines)
