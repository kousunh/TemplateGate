from __future__ import annotations

from ..core.model import CheckResult


def _cell(value) -> str:
    text = repr(value).replace("|", "\\|").replace("\n", " ")
    return text if len(text) <= 60 else text[:60] + "..."


def render_markdown(result: CheckResult) -> str:
    status = "✅ PASS" if result.passed else "❌ FAIL"
    lines = [
        f"## OfficeCheck: {status}",
        "",
        f"- Baseline: `{result.baseline}`",
        f"- Candidate: `{result.candidate}`",
        f"- Changes: {len(result.changes)} total / "
        f"{len(result.allowed)} allowed / {len(result.violations)} violations",
    ]
    if result.violations:
        lines += [
            "",
            "| Severity | Location | Attribute | Old | New | Rule |",
            "|---|---|---|---|---|---|",
        ]
        for v in result.violations:
            lines.append(
                f"| {v.severity} | `{v.change.location}` | {v.change.attribute} "
                f"| {_cell(v.change.old)} | {_cell(v.change.new)} | {v.rule} |"
            )
    if result.semantic_mode != "off" and result.semantic_findings:
        lines += ["", f"### Semantic checks ({result.semantic_mode})", ""]
        for f in result.semantic_findings:
            lines.append(f"- **{f.verdict}** — {f.check}: {f.message}")
    return "\n".join(lines)
