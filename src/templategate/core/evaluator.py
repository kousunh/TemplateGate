"""Default-deny evaluation of observed changes against a trusted policy.

Order of precedence per change:
  1. structural category set to "ignore"  -> dropped (not even reported)
  2. a recalculated formula result, under "recalculation: ignore" (the
     default)                             -> dropped, counted in the report
  3. matches a protect rule               -> violation ("protected")
  4. matches an allow rule                -> allowed
  5. otherwise                            -> violation ("not_allowed")

Violations are errors, except under ``mode: review_only``, where they are
reported as warnings so the run still passes.
"""

from __future__ import annotations

from .model import (
    Change,
    SEVERITY_ERROR,
    SEVERITY_WARNING,
    STRUCTURAL_ATTRIBUTES,
    Violation,
)
from .policy import MODE_REVIEW_ONLY, Policy
from .selector import match_attributes, match_selector

# The policy parser checks rules against this same map, so a structural key
# that evaluates to nothing here is rejected there rather than ignored.
_STRUCTURAL_ATTRS = STRUCTURAL_ATTRIBUTES


def evaluate(changes: list[Change], policy: Policy) -> tuple[list[Change], list[Violation]]:
    """Return (allowed_changes, violations)."""
    ignored_attrs = {
        attr
        for key, attr in _STRUCTURAL_ATTRS.items()
        if policy.structural_setting(key) == "ignore"
    }
    severity = SEVERITY_WARNING if policy.mode == MODE_REVIEW_ONLY else SEVERITY_ERROR

    allowed: list[Change] = []
    violations: list[Violation] = []
    ignore_recalculated = policy.recalculation == "ignore"

    for change in changes:
        if change.attribute in ignored_attrs:
            continue
        # The formula is byte-identical and only its stored answer moved, so
        # nothing about the document's logic changed — Excel recomputed on
        # save, or a library discarded the cache.  Dropped before protect, on
        # purpose: `protect: [value]` on a computed cell would otherwise fail
        # every real-Excel save, and the formula that produces the answer is
        # still guarded under `formula`.  Set `recalculation: strict` to judge
        # these like any other value change.
        if ignore_recalculated and change.recalculated:
            continue

        protected = any(
            match_selector(rule.selector, change.location)
            and match_attributes(rule.attributes, change.attribute)
            for rule in policy.protect
        )
        if protected:
            violations.append(
                Violation(
                    change=change,
                    rule="protected",
                    severity=severity,
                    message=f"protected attribute '{change.attribute}' changed",
                )
            )
            continue

        if any(
            match_selector(rule.selector, change.location)
            and match_attributes(rule.attributes, change.attribute)
            for rule in policy.allow
        ):
            allowed.append(change)
            continue

        violations.append(
            Violation(
                change=change,
                rule="not_allowed",
                severity=severity,
                message=(
                    f"change to '{change.attribute}' is not allowed "
                    "by the policy (default deny)"
                ),
            )
        )
    return allowed, violations
