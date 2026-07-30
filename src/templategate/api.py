"""Public API: snapshot / diff / check."""

from __future__ import annotations

import xml.etree.ElementTree as ElementTree
import zipfile
from pathlib import Path

from .core.evaluator import evaluate
from .core.model import Change, CheckResult, SEVERITY_ERROR
from .core.policy import MODE_PAGE_EXTENSION, MODE_REVIEW_ONLY, Policy, load_policy

EXCEL_SUFFIXES = (".xlsx", ".xlsm")
WORD_SUFFIXES = (".docx",)


class DocumentError(ValueError):
    """A document could not be read: missing, corrupt or not an Office file."""


def _read_errors() -> tuple[type[BaseException], ...]:
    """Exceptions that mean "this file is not a readable Office document"."""
    errors: tuple[type[BaseException], ...] = (
        zipfile.BadZipFile, OSError, ElementTree.ParseError, KeyError,
    )
    try:
        from openpyxl.utils.exceptions import InvalidFileException

        errors += (InvalidFileException,)
    except ImportError:  # pragma: no cover - openpyxl is a hard dependency
        pass
    try:
        from docx.opc.exceptions import OpcError

        errors += (OpcError,)
    except ImportError:  # pragma: no cover - python-docx is a hard dependency
        pass
    try:
        from lxml.etree import XMLSyntaxError

        errors += (XMLSyntaxError,)
    except ImportError:  # pragma: no cover - lxml ships with python-docx
        pass
    return errors


READ_ERRORS = _read_errors()


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
    try:
        return take_snapshot(path)
    except READ_ERRORS as exc:
        raise DocumentError(f"cannot read {path}: {exc}") from exc


def diff(baseline: str | Path, candidate: str | Path, *,
         align: bool = False) -> list[Change]:
    target = detect_target(baseline)
    if detect_target(candidate) != target:
        raise ValueError("baseline and candidate must be the same document type")
    if target == "excel":
        from .excel.diff import diff_snapshots

        return diff_snapshots(snapshot(baseline), snapshot(candidate))
    from .word.diff import diff_snapshots

    return diff_snapshots(snapshot(baseline), snapshot(candidate), align=align)


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

        changes = diff_snapshots(base_snap, cand_snap)
    else:
        from .word.diff import diff_snapshots

        changes = diff_snapshots(base_snap, cand_snap,
                                 align=policy.mode == MODE_PAGE_EXTENSION)
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

    # Under review_only the evaluator already downgrades violations to
    # warnings, so nothing here can fail the run.
    passed = not any(v.severity == SEVERITY_ERROR for v in violations)
    if (policy.mode != MODE_REVIEW_ONLY and mode == "gate"
            and any(f.verdict in ("fail", "error") for f in findings)):
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
