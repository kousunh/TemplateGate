"""Public API: snapshot / diff / check."""

from __future__ import annotations

from pathlib import Path

from .core.evaluator import evaluate
from .core.model import Change, CheckResult, SEVERITY_ERROR
from .core.policy import Policy, load_policy

EXCEL_SUFFIXES = (".xlsx", ".xlsm")
WORD_SUFFIXES = (".docx",)


def detect_target(path: str | Path) -> str:
    suffix = Path(path).suffix.lower()
    if suffix in EXCEL_SUFFIXES:
        return "excel"
    if suffix in WORD_SUFFIXES:
        return "word"
    raise ValueError(f"unsupported file type: {path} (expected .xlsx/.xlsm/.docx)")


def snapshot(path: str | Path) -> dict:
    if detect_target(path) == "excel":
        from .excel.snapshot import take_snapshot
    else:
        from .word.snapshot import take_snapshot
    return take_snapshot(path)


def diff(baseline: str | Path, candidate: str | Path) -> list[Change]:
    target = detect_target(baseline)
    if detect_target(candidate) != target:
        raise ValueError("baseline and candidate must be the same document type")
    if target == "excel":
        from .excel.diff import diff_snapshots
    else:
        from .word.diff import diff_snapshots
    return diff_snapshots(snapshot(baseline), snapshot(candidate))


def check(
    baseline: str | Path,
    candidate: str | Path,
    policy: Policy | str | Path,
    *,
    semantic_mode: str | None = None,
) -> CheckResult:
    if not isinstance(policy, Policy):
        policy = load_policy(policy)

    target = detect_target(baseline)
    if detect_target(candidate) != target:
        raise ValueError("baseline and candidate must be the same document type")
    if policy.target not in ("auto", target):
        raise ValueError(
            f"policy targets {policy.target!r} but the documents are {target!r}"
        )

    base_snap = snapshot(baseline)
    cand_snap = snapshot(candidate)
    if target == "excel":
        from .excel.diff import diff_snapshots
    else:
        from .word.diff import diff_snapshots
    changes = diff_snapshots(base_snap, cand_snap)
    allowed, violations = evaluate(changes, policy)

    mode = semantic_mode if semantic_mode is not None else policy.semantic.mode
    findings = []
    if mode != "off":
        from .semantic.base import extract_text, get_provider

        provider = get_provider(policy.semantic.provider)
        findings = provider.run_checks(
            policy.semantic,
            extract_text(base_snap),
            extract_text(cand_snap),
        )

    passed = not any(v.severity == SEVERITY_ERROR for v in violations)
    if mode == "gate" and any(f.verdict in ("fail", "error") for f in findings):
        passed = False

    return CheckResult(
        passed=passed,
        target=target,
        baseline=str(baseline),
        candidate=str(candidate),
        changes=changes,
        allowed=allowed,
        violations=violations,
        semantic_mode=mode,
        semantic_findings=findings,
        meta={"policy": policy.source_path, "policy_mode": policy.mode},
    )
