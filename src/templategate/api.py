"""Public API: snapshot / diff / check."""

from __future__ import annotations

from pathlib import Path

from .core.evaluator import evaluate
from .core.model import Change, CheckResult, SEVERITY_ERROR
from .core.policy import MODE_PAGE_EXTENSION, MODE_REVIEW_ONLY, Policy, load_policy

EXCEL_SUFFIXES = (".xlsx", ".xlsm")
WORD_SUFFIXES = (".docx",)


class DocumentError(ValueError):
    """A document could not be read: missing, corrupt or not an Office file."""


def detect_target(path: str | Path) -> str:
    suffix = Path(path).suffix.lower()
    if suffix in EXCEL_SUFFIXES:
        return "excel"
    if suffix in WORD_SUFFIXES:
        return "word"
    raise ValueError(f"unsupported file type: {path} (expected .xlsx/.xlsm/.docx)")


def snapshot(path: str | Path) -> dict:
    target = detect_target(path)
    from .core.package import package_problem

    # An ambiguous container is refused outright rather than degraded: there
    # is no "best effort" reading of a file whose contents depend on which
    # reader opens it.
    problem = package_problem(path)
    if problem is not None:
        raise DocumentError(f"cannot read {path}: {problem}")
    if target == "excel":
        from .excel.snapshot import take_snapshot
    else:
        from .word.snapshot import take_snapshot
    try:
        return take_snapshot(path)
    except DocumentError:
        raise
    except Exception as exc:
        # Deliberately broad.  A damaged package can make any parser raise
        # almost anything — a missing styles part surfaces as IndexError, a
        # truncated one as a parse error — and a traceback is the one outcome
        # a gate must never produce.  The exception type travels into the
        # report, so a bug here still reads as a bug rather than as damage.
        return _degraded_snapshot(path, target, exc)


def _degraded_snapshot(path: str | Path, target: str, exc: BaseException) -> dict:
    """What can still be learned about a document the library cannot open.

    A candidate that lost a part it still references — a deleted footer, a
    deleted numbering definition — is *damaged*, not unreadable: the zip opens
    fine and every surviving part can be compared.  Reporting that as a tool
    error would file real damage under "something went wrong on our side" and
    exit 2, so instead the package layer carries on alone and the check fails.
    Only a container that will not open at all is an error.
    """
    from .core.package import list_part_names, take_package_snapshot

    try:
        part_names = list_part_names(path)
        package = take_package_snapshot(path)
    except Exception:
        raise DocumentError(f"cannot read {path}: {exc}") from exc
    return {
        "target": target,
        "degraded": f"{type(exc).__name__}: {exc}",
        "package": package,
        "part_names": part_names,
    }


def _diff_snapshots(base: dict, cand: dict, target: str, *, align: bool) -> list[Change]:
    """Compare two snapshots, degrading to the package layer when one is partial."""
    if base.get("degraded") or cand.get("degraded"):
        from .core.package import diff_packages, diff_part_names

        # The raw name list names every part that came or went, in plain
        # language; the inventory adds the ones whose contents changed.  Where
        # both have something to say about a part, the plain wording wins.
        by_name = diff_part_names(base, cand)
        named = {change.location for change in by_name}
        return by_name + [c for c in diff_packages(base, cand)
                          if c.location not in named]
    if target == "excel":
        from .excel.diff import diff_snapshots

        return diff_snapshots(base, cand)
    from .word.diff import diff_snapshots

    return diff_snapshots(base, cand, align=align)


def diff(baseline: str | Path, candidate: str | Path, *,
         align: bool = False) -> list[Change]:
    target = detect_target(baseline)
    if detect_target(candidate) != target:
        raise ValueError("baseline and candidate must be the same document type")
    return _diff_snapshots(snapshot(baseline), snapshot(candidate), target,
                           align=align)


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
    changes = _diff_snapshots(base_snap, cand_snap, target,
                              align=policy.mode == MODE_PAGE_EXTENSION)
    allowed, violations = evaluate(changes, policy)

    from .core.liveness import policy_warnings

    warnings = policy_warnings(policy, base_snap)

    degraded = {
        role: snap["degraded"]
        for role, snap in (("baseline", base_snap), ("candidate", cand_snap))
        if snap.get("degraded")
    }

    mode = semantic_mode if semantic_mode is not None else policy.semantic.mode
    findings = []
    if mode != "off" and not degraded:
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
    # A document that would not open was never verified, so it cannot pass —
    # not even under review_only, and not even if a structural "ignore"
    # happened to silence every part that went missing.
    if degraded:
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
        warnings=warnings,
        meta={"policy": policy.source_path, "policy_mode": policy.mode,
              "degraded": degraded},
    )
