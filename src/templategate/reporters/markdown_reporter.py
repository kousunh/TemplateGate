from __future__ import annotations

from ..core.messages import Translator, say
from ..core.model import CheckResult
from .text_reporter import _grouped, _span, _worth_showing, human, verdict_of


def _escape(text: str) -> str:
    """Keep a value from breaking out of its Markdown table cell."""
    return str(text).replace("|", "\\|").replace("\n", " ")


def _cell(value, t: Translator) -> str:
    return _escape(human(value, 60, t))


def render_markdown(result: CheckResult, lang: str = "en") -> str:
    t = Translator(lang)
    status = t("report.md.pass") if result.passed else t("report.md.fail")
    lines = [
        t("report.md.title", status=status),
        "",
        f"- {t('report.md.baseline')}: `{result.baseline}`",
        f"- {t('report.md.candidate')}: `{result.candidate}`",
        "- " + t("report.md.summary", total=len(result.changes),
                 allowed=len(result.allowed),
                 violations=len(result.violations)),
    ]
    if result.recalculated_ignored:
        lines.append(t("report.md.recalculated",
                       count=result.recalculated_ignored))
    for role, reason in (result.meta.get("degraded") or {}).items():
        lines += ["",
                  t("damaged.md.line", role=t(f"role.{role}"),
                    reason=_escape(reason)),
                  t("damaged.md.parts_only")]
    if result.warnings:
        lines += ["", t("report.md.warnings"), ""]
        lines += [f"- {_escape(say(warning, t))}" for warning in result.warnings]
    if result.violations:
        lines += [""]
        if result.meta.get("policy_mode") == "review_only":
            lines += [t("report.md.review_only"), ""]
        lines += [t("report.md.columns"), "|---|---|---|---|---|---|"]
        for group, members in _grouped(result.violations):
            if group:
                first = members[0].change
                lines.append(
                    f"| {t('severity.' + members[0].severity)} | `{_escape(_span(members))}` "
                    f"| {_escape(first.attribute)} "
                    f"| {_escape(say(members[0].change.group, t))} "
                    f"| {t('report.md.cascade', count=len(members))} "
                    f"| {_escape(members[0].rule)} |"
                )
                continue
            v = members[0]
            if _worth_showing(v.change):
                before, after = _cell(v.change.old, t), _cell(v.change.new, t)
            else:
                # There is no Old/New worth printing, so the row carries the
                # explanation instead of two columns of "none".
                before = _escape(say(v.change.detail, t) or verdict_of(v, t))
                after = ""
            lines.append(
                f"| {t('severity.' + v.severity)} | `{_escape(v.change.location)}` "
                f"| {_escape(v.change.attribute)} "
                f"| {before} | {after} "
                f"| {_escape(v.rule)} |"
            )
    if result.semantic_mode != "off" and result.semantic_findings:
        lines += ["", t("report.md.semantic", mode=result.semantic_mode), ""]
        for f in result.semantic_findings:
            lines.append(f"- **{f.verdict}** — {f.check}: {f.message}")
    return "\n".join(lines)
