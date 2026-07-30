"""Reports have to be readable by whoever approves the change.

A violation that says "None -> None" or lists fifty-six shifted paragraphs
tells a reviewer nothing they can act on, so these tests pin what the report
actually says, not just that something was reported.
"""

import json
import zipfile

import pytest
from docx import Document
from openpyxl import Workbook

from templategate import check, diff
from templategate.core.policy import parse_policy
from templategate.core.selector import match_selector, quote_sheet
from templategate.reporters import render_json, render_markdown, render_text

STRICT_WORD = parse_policy({
    "target": "word",
    "protect": [{"selector": "*", "attributes": ["*"]}],
})


def _only(changes, attribute):
    return [c for c in changes if c.attribute == attribute]


# --- C1: a format change says which property changed ---------------------

def test_number_format_change_names_the_property(tmp_path):
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    for path, number_format in ((baseline, "#,##0"), (candidate, "#,##0,")):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A4"] = 1000
        ws["A4"].number_format = number_format
        wb.save(path)

    change = _only(diff(baseline, candidate), "format")[0]
    assert change.detail == "cell format changed: numfmt #,##0 -> #,##0,"
    assert change.old == {"numfmt": "#,##0"}
    assert change.new == {"numfmt": "#,##0,"}


def test_format_delta_carries_only_the_changed_property(tmp_path):
    """The reader wants the one thing that moved, not the whole style."""
    from openpyxl.styles import Font

    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    for path, color in ((baseline, "FF000000"), (candidate, "FFFFFFFF")):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "total"
        ws["A1"].font = Font(bold=True, color=color)
        wb.save(path)

    change = _only(diff(baseline, candidate), "format")[0]
    assert list(change.old) == ["font.color"]
    assert "font.color" in change.detail


def test_word_hidden_text_detail_names_the_property(tmp_path):
    from generate import rewrite_zip

    def build(name, run_properties):
        plain = tmp_path / f"{name}_plain.docx"
        document = Document()
        document.add_paragraph("x")
        document.save(str(plain))
        with zipfile.ZipFile(plain) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        body = (f"<w:p><w:r>{run_properties}<w:t>Liability is capped.</w:t>"
                "</w:r></w:p>")
        built = tmp_path / f"{name}.docx"
        rewrite_zip(plain, built, add={
            "word/document.xml": xml.replace("</w:body>", body + "</w:body>").encode()})
        return built

    baseline = build("base", "")
    candidate = build("cand", "<w:rPr><w:vanish/></w:rPr>")
    change = _only(diff(baseline, candidate), "format")[0]
    assert "vanish" in change.detail


# --- C2: a sheet name containing "!" is not the same as a cell reference --

def test_quote_sheet_only_quotes_what_needs_it():
    assert quote_sheet("Sheet1") == "Sheet1"
    assert quote_sheet("計画表") == "計画表"
    assert quote_sheet("Q1!Q4") == "'Q1!Q4'"
    assert quote_sheet("Q1#Draft") == "'Q1#Draft'"
    assert quote_sheet("It's") == "'It''s'"


def test_cell_selector_never_reaches_a_same_named_sheet():
    """"Q1!Q4" is cell Q4 of sheet Q1, and must stay only that."""
    assert match_selector("Q1!Q4", "Q1!Q4")
    assert not match_selector("Q1!Q4", "'Q1!Q4'!A1")
    assert not match_selector("Q1!Q4", "'Q1!Q4'!Q4")


def test_quoted_sheet_selector_never_reaches_the_cell():
    assert match_selector("'Q1!Q4'!A1", "'Q1!Q4'!A1")
    assert match_selector("'Q1!Q4'", "'Q1!Q4'!A1")
    assert not match_selector("'Q1!Q4'", "Q1!Q4")
    assert not match_selector("'Q1!Q4'!A1:B9", "Q1!Q4")


def test_locations_quote_ambiguous_sheet_names(tmp_path):
    """A workbook with both readings present must keep them apart."""
    def build(path, secret):
        wb = Workbook()
        ws = wb.active
        ws.title = "Q1"
        ws["Q4"] = secret
        wb.create_sheet("Q1!Q4")["A1"] = "scratch"
        wb.save(path)

    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    build(baseline, "PROTECTED secret")
    build(candidate, "LEAKED")

    locations = {c.location for c in diff(baseline, candidate)}
    assert locations == {"Q1!Q4"}

    # A rule written for the scratch *sheet* must not touch the secret cell.
    policy = parse_policy({
        "target": "excel",
        "allow": [{"selector": "'Q1!Q4'", "attributes": ["value"]}],
    })
    assert not check(baseline, candidate, policy).passed


def test_edits_on_the_ambiguous_sheet_are_addressable(tmp_path):
    def build(path, scratch):
        wb = Workbook()
        ws = wb.active
        ws.title = "Q1"
        ws["Q4"] = "secret"
        wb.create_sheet("Q1!Q4")["A1"] = scratch
        wb.save(path)

    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    build(baseline, "scratch")
    build(candidate, "rewritten")

    assert {c.location for c in diff(baseline, candidate)} == {"'Q1!Q4'!A1"}
    policy = parse_policy({
        "target": "excel",
        "allow": [{"selector": "'Q1!Q4'", "attributes": ["value"]}],
    })
    assert check(baseline, candidate, policy).passed


# --- C3: one edit reads as one edit --------------------------------------

def _numbered(tmp_path, name, skip=None):
    path = tmp_path / f"{name}.docx"
    document = Document()
    for number in range(1, 21):
        if number == skip:
            continue
        document.add_paragraph(f"Clause {number}.")
    document.save(str(path))
    return path


def test_one_deletion_still_reports_every_change(tmp_path):
    """Collapsing is a reporting choice; the policy still sees everything."""
    baseline = _numbered(tmp_path, "base")
    candidate = _numbered(tmp_path, "cand", skip=5)
    changes = diff(baseline, candidate)
    assert len(changes) > 10
    result = check(baseline, candidate, STRICT_WORD)
    assert not result.passed
    assert len(result.violations) == len(changes)


def test_shifted_paragraphs_are_grouped_under_the_real_edit(tmp_path):
    baseline = _numbered(tmp_path, "base")
    candidate = _numbered(tmp_path, "cand", skip=5)
    groups = {c.group for c in diff(baseline, candidate) if c.group}
    assert groups == {"1 paragraph removed at p5"}


def test_text_report_collapses_the_cascade(tmp_path):
    baseline = _numbered(tmp_path, "base")
    candidate = _numbered(tmp_path, "cand", skip=5)
    report = render_text(check(baseline, candidate, STRICT_WORD))
    violation_lines = [line for line in report.splitlines()
                       if line.startswith("  [error]")]
    assert len(violation_lines) <= 4, report
    assert any("content shifted because 1 paragraph removed at p5" in line
               for line in report.splitlines())


def test_json_report_keeps_every_change(tmp_path):
    """Machines get the full list even when humans get the summary."""
    baseline = _numbered(tmp_path, "base")
    candidate = _numbered(tmp_path, "cand", skip=5)
    result = check(baseline, candidate, STRICT_WORD)
    data = json.loads(render_json(result))
    assert len(data["violations"]) == len(result.violations)
    assert any(v["change"]["group"] for v in data["violations"])


def test_markdown_report_collapses_the_cascade(tmp_path):
    baseline = _numbered(tmp_path, "base")
    candidate = _numbered(tmp_path, "cand", skip=5)
    report = render_markdown(check(baseline, candidate, STRICT_WORD))
    rows = [line for line in report.splitlines()
            if line.startswith("| error")]
    assert len(rows) <= 4, report


def _reordered(tmp_path, name, order):
    path = tmp_path / f"{name}.docx"
    document = Document()
    document.add_paragraph("Obligations")
    for text in order:
        document.add_paragraph(text)
    document.save(str(path))
    return path


CLAUSES = ["1. The Buyer shall pay within 14 days.",
           "2. The Seller shall deliver on time.",
           "3. The Supplier may suspend delivery while payment is overdue."]


def test_reordering_is_reported_as_a_move(tmp_path):
    baseline = _reordered(tmp_path, "base", CLAUSES)
    candidate = _reordered(tmp_path, "cand", [CLAUSES[2], CLAUSES[1], CLAUSES[0]])
    moves = _only(diff(baseline, candidate, align=True), "moved")
    assert moves
    assert all("moved from" in m.detail for m in moves)
    assert all(m.old != m.new for m in moves)


def test_page_extension_no_longer_waves_a_reorder_through(tmp_path):
    """Swapping two clauses is not a text edit, so a text rule cannot allow it."""
    baseline = _reordered(tmp_path, "base", CLAUSES)
    candidate = _reordered(tmp_path, "cand", [CLAUSES[2], CLAUSES[1], CLAUSES[0]])
    policy = parse_policy({
        "target": "word",
        "mode": "page_extension",
        "allow": [{"selector": "body", "attributes": ["text"]}],
    })
    assert not check(baseline, candidate, policy).passed


def test_page_extension_still_allows_a_plain_insertion(tmp_path):
    """The relaxation has to keep working: an insertion is not a move."""
    baseline = _reordered(tmp_path, "base", CLAUSES)
    candidate = _reordered(tmp_path, "cand",
                           [CLAUSES[0], "1a. Interest accrues after 14 days.",
                            CLAUSES[1], CLAUSES[2]])
    changes = diff(baseline, candidate, align=True)
    assert not _only(changes, "moved")
    assert [c.detail for c in changes] == ["paragraph added"]


# --- C4: wording a reader can trust --------------------------------------

def test_merged_cells_say_so(tmp_path):
    def build(path, merge):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"], ws["B1"] = "Q1", "Q2"
        if merge:
            ws.merge_cells("A1:B1")
        wb.save(path)

    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    build(baseline, False)
    build(candidate, True)
    merged = _only(diff(baseline, candidate), "merge")
    assert merged and "cells merged" in merged[0].detail

    unmerged = _only(diff(candidate, baseline), "merge")
    assert unmerged and "cells unmerged" in unmerged[0].detail


def test_macro_enabled_workbook_saved_as_xlsx_is_reported(tmp_path):
    from generate import rewrite_zip

    baseline = tmp_path / "base.xlsm"
    candidate = tmp_path / "cand.xlsx"
    wb = Workbook()
    wb.active["A1"] = "x"
    wb.save(baseline)
    rewrite_zip(baseline, candidate, add={})

    change = [c for c in diff(baseline, candidate)
              if c.location == "workbook#format"][0]
    assert (change.old, change.new) == ("xlsm", "xlsx")
    assert "loses its VBA project" in change.detail


def test_page_break_removal_says_page_break(tmp_path):
    from generate import rewrite_zip

    def build(name, body):
        plain = tmp_path / f"{name}_plain.docx"
        document = Document()
        document.add_paragraph("x")
        document.save(str(plain))
        with zipfile.ZipFile(plain) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        built = tmp_path / f"{name}.docx"
        rewrite_zip(plain, built, add={
            "word/document.xml": xml.replace("</w:body>", body + "</w:body>").encode()})
        return built

    baseline = build("base", '<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    candidate = build("cand", "<w:p><w:r></w:r></w:p>")
    breaks = [c for c in diff(baseline, candidate)
              if "page break" in c.detail]
    assert breaks
    assert "page break removed" in breaks[0].detail
