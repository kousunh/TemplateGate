"""Data model shared across TemplateGate.

A *snapshot* is a normalized dict extracted from a document.  Comparing two
snapshots yields ``Change`` records; evaluating changes against a policy
yields ``Violation`` records collected into a ``CheckResult``.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any

# Attribute names used across extractors and policies.
ATTR_VALUE = "value"
ATTR_FORMULA = "formula"
ATTR_FORMAT = "format"
ATTR_MERGE = "merge"
ATTR_CONDITIONAL_FORMATTING = "conditional_formatting"
ATTR_DATA_VALIDATION = "data_validation"
ATTR_DEFINED_NAMES = "defined_names"
ATTR_SHEET_STRUCTURE = "sheet_structure"
ATTR_IMAGES = "images"
ATTR_HEADER_FOOTER = "header_footer"
ATTR_PRINT_SETTINGS = "print_settings"
ATTR_VBA = "vba"
ATTR_TEXT = "text"
ATTR_STYLE = "style"
ATTR_SECTION = "section"
ATTR_TABLE = "table"

# OOXML package parts that editing libraries silently drop on save.
ATTR_CHARTS = "charts"
ATTR_PIVOT_TABLES = "pivot_tables"
ATTR_DRAWINGS = "drawings"
ATTR_COMMENTS = "comments"
ATTR_EMBEDDED = "embedded"
ATTR_CUSTOM_XML = "custom_xml"
# Every other package part, and the external targets of relationships.
ATTR_PARTS = "parts"
ATTR_LINKS = "links"

# Surfaces that decide what a reader can see or change, rather than what the
# document says.
ATTR_PROTECTION = "protection"
ATTR_SHEET_SETTINGS = "sheet_settings"
ATTR_LAYOUT = "layout"

# Word surfaces that live below the text: how a paragraph is laid out, what a
# field computes, where a cross-reference points, and whether an edit was
# recorded as a tracked change.
ATTR_PARAGRAPH_FORMAT = "paragraph_format"
ATTR_FIELD = "field"
ATTR_BOOKMARK = "bookmark"
ATTR_REVISION = "revision"
ATTR_CONTENT_CONTROL = "content_control"
# A block whose content survived but whose position did not.
ATTR_MOVED = "moved"
# Markup in a block that no other attribute accounts for.  The backstop that
# keeps an unmodelled feature from being an invisible one.
ATTR_MARKUP = "markup"

# Every attribute a policy may name in an allow or protect rule.  Kept here
# so the parser can check a rule against the same list the evaluator uses —
# a policy that names something no attribute is called protects nothing.
ALL_ATTRIBUTES = frozenset({
    ATTR_VALUE, ATTR_FORMULA, ATTR_FORMAT, ATTR_MERGE,
    ATTR_CONDITIONAL_FORMATTING, ATTR_DATA_VALIDATION, ATTR_DEFINED_NAMES,
    ATTR_SHEET_STRUCTURE, ATTR_IMAGES, ATTR_HEADER_FOOTER, ATTR_PRINT_SETTINGS,
    ATTR_VBA, ATTR_TEXT, ATTR_STYLE, ATTR_SECTION, ATTR_TABLE,
    ATTR_CHARTS, ATTR_PIVOT_TABLES, ATTR_DRAWINGS, ATTR_COMMENTS,
    ATTR_EMBEDDED, ATTR_CUSTOM_XML, ATTR_PARTS, ATTR_LINKS,
    ATTR_PROTECTION, ATTR_SHEET_SETTINGS, ATTR_LAYOUT,
    ATTR_PARAGRAPH_FORMAT, ATTR_FIELD, ATTR_BOOKMARK, ATTR_REVISION,
    ATTR_CONTENT_CONTROL, ATTR_MOVED, ATTR_MARKUP,
})

# structural policy key -> the change attribute it governs.
STRUCTURAL_ATTRIBUTES = {
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

SEVERITY_ERROR = "error"
SEVERITY_WARNING = "warning"


@dataclass
class Change:
    """A single observed difference between baseline and candidate."""

    location: str  # e.g. "Sheet1!B2", "sheet:Sheet2", "p12", "table1!r2c3"
    attribute: str  # one of the ATTR_* constants
    old: Any = None
    new: Any = None
    detail: str = ""
    # Changes that are all knock-on effects of one edit share a group, so a
    # report can say "one paragraph removed" instead of listing every
    # paragraph that shifted up behind it.  Evaluation ignores this: each
    # change is still judged on its own.
    group: str = ""
    # A value that moved while the formula producing it stayed byte-identical:
    # the answer was recomputed, not rewritten.  Set only by the comparison
    # layer; the policy decides what to do with it (see ``recalculation``).
    recalculated: bool = False
    def to_dict(self) -> dict:
        # ``detail`` and ``group`` may be Detail instances, which are strings
        # carrying the message id that produced them.  They serialise as the
        # English text they already are, so the JSON contract is untouched.
        return asdict(self)


@dataclass
class Violation:
    change: Change
    rule: str  # "protected" | "not_allowed" | "structural"
    severity: str = SEVERITY_ERROR
    message: str = ""

    def to_dict(self) -> dict:
        return {
            "change": self.change.to_dict(),
            "rule": self.rule,
            "severity": self.severity,
            "message": self.message,
        }


@dataclass
class SemanticFinding:
    check: str
    verdict: str  # "pass" | "fail" | "warning" | "error"
    message: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class CheckResult:
    passed: bool
    target: str  # "excel" | "word"
    baseline: str
    candidate: str
    changes: list[Change] = field(default_factory=list)
    allowed: list[Change] = field(default_factory=list)
    violations: list[Violation] = field(default_factory=list)
    semantic_mode: str = "off"
    semantic_findings: list[SemanticFinding] = field(default_factory=list)
    # Things wrong with the policy itself rather than with the document.
    warnings: list[str] = field(default_factory=list)
    # Cached formula results that moved while their formulas did not, and
    # that ``recalculation: ignore`` dropped before evaluation.  Counted
    # rather than listed: they are noise the reader asked not to see, but
    # silently losing several changes from the totals would be worse.
    recalculated_ignored: int = 0
    meta: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "tool": "templategate",
            "passed": self.passed,
            "target": self.target,
            "baseline": self.baseline,
            "candidate": self.candidate,
            "summary": {
                "total_changes": len(self.changes),
                "allowed": len(self.allowed),
                "violations": len(self.violations),
                "errors": sum(1 for v in self.violations if v.severity == SEVERITY_ERROR),
                "warnings": sum(1 for v in self.violations if v.severity == SEVERITY_WARNING),
                "recalculated_ignored": self.recalculated_ignored,
            },
            "warnings": list(self.warnings),
            "violations": [v.to_dict() for v in self.violations],
            "allowed_changes": [c.to_dict() for c in self.allowed],
            "semantic": {
                "mode": self.semantic_mode,
                "findings": [f.to_dict() for f in self.semantic_findings],
            },
            "meta": self.meta,
        }
