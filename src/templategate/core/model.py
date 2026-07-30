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

    def to_dict(self) -> dict:
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
            },
            "violations": [v.to_dict() for v in self.violations],
            "allowed_changes": [c.to_dict() for c in self.allowed],
            "semantic": {
                "mode": self.semantic_mode,
                "findings": [f.to_dict() for f in self.semantic_findings],
            },
            "meta": self.meta,
        }
