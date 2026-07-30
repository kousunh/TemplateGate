"""Damaged documents are a FAIL, not a tool error.

A candidate that has lost a part it still references still opens as a zip:
every surviving part can be compared and the missing one can be named.
Reporting that as an execution error would file real damage under "something
went wrong on our side" and hand back exit 2, which CI reads as "the gate
broke", not "the document broke".  Exit 2 stays for containers that will not
open at all.
"""

import json

import pytest

from templategate import check, diff, snapshot
from templategate.api import DocumentError
from templategate.cli import main
from templategate.core.policy import parse_policy

STRICT_WORD = parse_policy({
    "target": "word",
    "protect": [{"selector": "*", "attributes": ["*"]}],
})


@pytest.fixture()
def damaged(fixtures, tmp_path):
    """A document missing a part that its own relationships still point at."""
    from generate import rewrite_zip

    baseline = tmp_path / "base.docx"
    candidate = tmp_path / "cand.docx"
    rewrite_zip(fixtures["word_baseline"], baseline, add={})
    rewrite_zip(baseline, candidate, drop=("word/styles.xml",))
    return baseline, candidate


def test_missing_part_is_reported_by_name(damaged):
    baseline, candidate = damaged
    changes = diff(baseline, candidate)
    missing = [c for c in changes if c.location.endswith("word/styles.xml")]
    assert missing, [c.location for c in changes]
    assert missing[0].new is None
    assert missing[0].detail == "word/styles.xml is missing from the candidate"


def test_damaged_candidate_fails_instead_of_erroring(damaged):
    baseline, candidate = damaged
    result = check(baseline, candidate, STRICT_WORD)
    assert not result.passed
    assert result.meta["degraded"]
    assert "candidate" in result.meta["degraded"]


def test_damaged_candidate_exits_one(damaged, tmp_path, capsys):
    baseline, candidate = damaged
    policy = tmp_path / "p.yaml"
    policy.write_text('version: 1\ntarget: word\nprotect:\n  - selector: "*"\n'
                      '    attributes: ["*"]\n', encoding="utf-8")
    code = main(["check", "--baseline", str(baseline), "--candidate", str(candidate),
                 "--policy", str(policy)])
    assert code == 1
    out = capsys.readouterr().out
    assert "word/styles.xml is missing from the candidate" in out
    assert "damaged" in out


def test_damaged_document_cannot_pass_under_review_only(damaged):
    """review_only downgrades violations, but an unverifiable file is not a pass."""
    baseline, candidate = damaged
    policy = parse_policy({"target": "word", "mode": "review_only",
                           "allow": [{"selector": "*", "attributes": ["*"]}]})
    assert not check(baseline, candidate, policy).passed


def test_damaged_document_cannot_pass_when_parts_are_ignored(damaged):
    """Silencing the parts category must not silence "it would not open"."""
    baseline, candidate = damaged
    policy = parse_policy({
        "target": "word",
        "allow": [{"selector": "*", "attributes": ["*"]}],
        "structural": {"parts": "ignore", "links": "ignore"},
    })
    assert not check(baseline, candidate, policy).passed


def test_degraded_snapshot_still_lists_parts(damaged):
    baseline, candidate = damaged
    snap = snapshot(candidate)
    assert snap["degraded"]
    assert "word/document.xml" in snap["part_names"]
    assert "word/styles.xml" not in snap["part_names"]
    json.dumps(snap)


def test_a_file_that_is_not_a_container_is_still_an_error(fixtures, tmp_path, capsys):
    """Exit 2 survives for input that is not a readable package at all."""
    broken = tmp_path / "broken.docx"
    broken.write_bytes(b"this is not a zip")
    with pytest.raises(DocumentError):
        snapshot(broken)

    policy = tmp_path / "p.yaml"
    policy.write_text("version: 1\ntarget: word\n", encoding="utf-8")
    code = main(["check", "--baseline", str(fixtures["word_baseline"]),
                 "--candidate", str(broken), "--policy", str(policy)])
    assert code == 2
    assert "cannot read" in capsys.readouterr().err


STRICT_EXCEL = parse_policy({
    "target": "excel",
    "protect": [{"selector": "*", "attributes": ["*"]}],
})


@pytest.mark.parametrize("part", ["xl/workbook.xml", "xl/styles.xml",
                                  "xl/_rels/workbook.xml.rels"])
def test_damaged_excel_workbook_fails_without_a_traceback(fixtures, tmp_path, part):
    """Damaged packages make parsers raise all sorts; none may escape.

    A missing styles part surfaces as IndexError rather than any recognisable
    "bad file" exception, which used to leave the CLI in a traceback.
    """
    from generate import rewrite_zip

    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    rewrite_zip(fixtures["excel_baseline"], baseline, add={})
    rewrite_zip(baseline, candidate, drop=(part,))

    result = check(baseline, candidate, STRICT_EXCEL)
    assert not result.passed
    assert result.meta["degraded"]
    assert any(part in v.change.location for v in result.violations)


def test_missing_worksheet_is_reported_as_a_lost_sheet(fixtures, tmp_path):
    """openpyxl tolerates a missing sheet part, so it stays a normal read.

    The sheet is then reported by name, which is more use than a part
    filename — and is why worksheet parts are left out of the inventory.
    """
    from generate import rewrite_zip

    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    rewrite_zip(fixtures["excel_baseline"], baseline, add={})
    rewrite_zip(baseline, candidate, drop=("xl/worksheets/sheet1.xml",))

    assert snapshot(candidate).get("degraded") is None
    changes = diff(baseline, candidate)
    assert ("sheet:計画表", "sheet removed") in {(c.location, c.detail) for c in changes}
    assert not check(baseline, candidate, STRICT_EXCEL).passed
