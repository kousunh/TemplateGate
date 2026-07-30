"""Default-deny evaluation of observed changes against a trusted policy.

Order of precedence per change:
  1. structural category set to "ignore"  -> dropped (not even reported)
  2. matches a protect rule               -> violation ("protected")
  3. matches an allow rule                -> allowed
  4. otherwise                            -> violation ("not_allowed")

Violations are errors, except under ``mode: review_only``, where they are
reported as warnings so the run still passes.
"""

from __future__ import annotations

from .model import (
    ATTR_CHARTS,
    ATTR_COMMENTS,
    ATTR_CUSTOM_XML,
    ATTR_DEFINED_NAMES,
    ATTR_DRAWINGS,
    ATTR_EMBEDDED,
    ATTR_IMAGES,
    ATTR_LINKS,
    ATTR_PARTS,
    ATTR_PIVOT_TABLES,
    ATTR_SHEET_STRUCTURE,
    ATTR_TABLE,
    Change,
    SEVERITY_ERROR,
    SEVERITY_WARNING,
    Violation,
)
from .policy import MODE_REVIEW_ONLY, Policy
from .selector import match_attributes, match_selector

# structural policy key -> change attribute it governs
_STRUCTURAL_ATTRS = {
    "sheets": ATTR_SHEET_STRUCTURE,
    "images": ATTR_IMAGES,
    "defined_names": ATTR_DEFINED_NAMES,
    "tables": ATTR_TABLE,
    "charts": ATTR_CHARTS,
    "pivot_tables": ATTR_PIVOT_TABLES,
    "drawings": ATTR_DRAWINGS,
    "comments": ATTR_COMMENTS,
    "embedded": ATTR_EMBEDDED,
    "custom_xml": ATTR_CUSTOM_XML,
    "parts": ATTR_PARTS,
    "links": ATTR_LINKS,
}


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
    for change in changes:
        if change.attribute in ignored_attrs:
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
                    message=f"protected attribute '{change.attribute}' changed at {change.location}",
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
                    f"change to '{change.attribute}' at {change.location} "
                    "is not allowed by the policy (default deny)"
                ),
            )
        )
    return allowed, violations
