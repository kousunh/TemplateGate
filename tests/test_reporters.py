"""Report rendering: escaping and review_only presentation."""

import json

from templategate.core.model import (
    SEVERITY_WARNING,
    Change,
    CheckResult,
    Violation,
)
from templategate.reporters import render_json, render_markdown, render_text


def _result(location="Sheet1!B2", *, policy_mode="normal_input", passed=False):
    change = Change(location, "value", old="a|b", new="c\nd")
    return CheckResult(
        passed=passed,
        target="excel",
        baseline="base.xlsx",
        candidate="cand.xlsx",
        changes=[change],
        violations=[Violation(change=change, rule="not_allowed",
                              severity=SEVERITY_WARNING, message="nope")],
        meta={"policy_mode": policy_mode},
    )


def _delimiters(row: str) -> int:
    """Count the cell delimiters, ignoring escaped pipes inside cells."""
    return row.replace("\\|", "").count("|")


def test_markdown_escapes_pipes_in_the_location():
    """A pipe in a sheet name must not break out of the table cell."""
    markdown = render_markdown(_result("Sales|Q1!B2"))
    row = [line for line in markdown.splitlines() if "Sales" in line][0]
    assert "`Sales\\|Q1!B2`" in row
    assert _delimiters(row) == 7  # six columns


def test_markdown_escapes_pipes_and_newlines_in_values():
    markdown = render_markdown(_result())
    row = [line for line in markdown.splitlines() if "not_allowed" in line][0]
    assert "\\|" in row
    assert "\n" not in row.strip("\n")
    assert _delimiters(row) == 7


def test_markdown_flags_review_only_as_non_blocking():
    markdown = render_markdown(_result(policy_mode="review_only", passed=True))
    assert "review_only" in markdown
    assert "not blocking" in markdown


def test_text_flags_review_only_as_non_blocking():
    text = render_text(_result(policy_mode="review_only", passed=True))
    assert "review_only" in text
    assert "not blocking" in text
    assert "TemplateGate: PASS" in text


def test_text_report_is_plain_for_normal_input():
    text = render_text(_result())
    assert "Violations:" in text
    assert "review_only" not in text


def test_json_counts_warnings_separately_from_errors():
    data = json.loads(render_json(_result(policy_mode="review_only", passed=True)))
    assert data["summary"] == {
        "total_changes": 1,
        "allowed": 0,
        "violations": 1,
        "errors": 0,
        "warnings": 1,
    }
