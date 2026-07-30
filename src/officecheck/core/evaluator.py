"""Default-deny evaluation of observed changes against a trusted policy.

Order of precedence per change:
  1. structural category set to "ignore"  -> dropped (not even reported)
  2. matches a protect rule               -> violation ("protected")
  3. matches an allow rule                -> allowed
  4. otherwise                            -> violation ("not_allowed")
"""

from __future__ import annotations

from .model import (
    ATTR_DEFINED_NAMES,
    ATTR_IMAGES,
    ATTR_SHEET_STRUCTURE,
    ATTR_TABLE,
    Change,
    SEVERITY_ERROR,
    Violation,
)
from .policy import Policy
from .selector import match_attributes, match_selector

# structural policy key -> change attribute it governs
_STRUCTURAL_ATTRS = {
    "sheets": ATTR_SHEET_STRUCTURE,
    "images": ATTR_IMAGES,
    "defined_names": ATTR_DEFINED_NAMES,
    "tables": ATTR_TABLE,
}


def evaluate(changes: list[Change], policy: Policy) -> tuple[list[Change], list[Violation]]:
    """Return (allowed_changes, violations)."""
    ignored_attrs = {
        attr
        for key, attr in _STRUCTURAL_ATTRS.items()
        if policy.structural_setting(key) == "ignore"
    }

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
                    severity=SEVERITY_ERROR,
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
                severity=SEVERITY_ERROR,
                message=(
                    f"change to '{change.attribute}' at {change.location} "
                    "is not allowed by the policy (default deny)"
                ),
            )
        )
    return allowed, violations
