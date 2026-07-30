from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from templategate import check, diff, snapshot
from templategate.core.model import (
    ATTR_CONDITIONAL_FORMATTING,
    ATTR_DATA_VALIDATION,
    ATTR_FORMULA,
    ATTR_VALUE,
)
from templategate.core.policy import parse_policy


def test_snapshot_contents(fixtures):
    snap = snapshot(fixtures["excel_baseline"])
    assert snap["target"] == "excel"
    assert set(snap["sheets"]) == {"計画表", "メモ"}
    sheet = snap["sheets"]["計画表"]
    assert sheet["cells"]["A1"]["value"] == "架空市 水道整備計画(サンプル)"
    assert sheet["cells"]["D8"]["formula"] == "=SUM(D4:D6)"
    assert "A1:D1" in sheet["merges"]
    assert sheet["conditional_formatting"]
    assert sheet["data_validation"]
    assert sheet["print"]["orientation"] == "landscape"
    assert sheet["header_footer"]


def test_identical_files_have_no_changes(fixtures):
    assert diff(fixtures["excel_baseline"], fixtures["excel_baseline"]) == []


def test_good_edit_passes(fixtures):
    result = check(fixtures["excel_baseline"], fixtures["excel_good"],
                   fixtures["excel_policy"])
    assert result.passed, [v.message for v in result.violations]
    assert {c.location for c in result.allowed} == {"計画表!B4", "計画表!B5"}
    assert result.semantic_mode == "off"


def test_bad_edit_fails_with_expected_violations(fixtures):
    result = check(fixtures["excel_baseline"], fixtures["excel_bad"],
                   fixtures["excel_policy"])
    assert not result.passed

    by_key = {(v.change.location, v.change.attribute): v for v in result.violations}
    # destroyed formula is caught by the protect rule
    assert by_key[("計画表!D8", "formula")].rule == "protected"
    # header format change is default-denied
    assert by_key[("計画表!A3", "format")].rule == "not_allowed"
    # deleted sheet is a structural violation
    assert ("sheet:メモ", "sheet_structure") in by_key
    # the allowed value edit still goes through
    assert any(c.location == "計画表!B4" for c in result.allowed)


def _sheet_with_values(path, title, values):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    for coord, value in values.items():
        ws[coord] = value
    wb.save(path)


def test_renamed_sheet_still_compares_cells(tmp_path):
    """A rename must not carry the sheet's contents out of sight."""
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    _sheet_with_values(baseline, "Plan", {"A1": "Budget", "B1": 1000, "B2": "=B1*2"})
    _sheet_with_values(candidate, "Plan ", {"A1": "HACKED", "B1": 999999, "B2": 0})

    changes = diff(baseline, candidate)
    by_key = {(c.location, c.attribute) for c in changes}
    assert ("Plan!A1", ATTR_VALUE) in by_key
    assert ("Plan!B1", ATTR_VALUE) in by_key
    assert ("Plan!B2", ATTR_FORMULA) in by_key


def test_sheet_rename_with_sheets_ignore_is_not_a_bypass(tmp_path):
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    _sheet_with_values(baseline, "Plan", {"A1": "Budget", "B1": 1000, "B2": "=B1*2"})
    _sheet_with_values(candidate, "Plan ", {"A1": "HACKED", "B1": 999999, "B2": 0})

    policy = parse_policy({
        "target": "excel",
        "protect": [{"selector": "*", "attributes": ["*"]}],
        "structural": {"sheets": "ignore"},
    })
    result = check(baseline, candidate, policy)
    assert not result.passed
    assert result.violations


def test_pure_rename_is_reported_as_a_rename(tmp_path):
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    values = {"A1": "Budget", "B1": 1000}
    _sheet_with_values(baseline, "Plan", values)
    _sheet_with_values(candidate, "Plan2", values)

    changes = diff(baseline, candidate)
    details = {c.location: c.detail for c in changes}
    assert details == {
        "sheet:Plan": "sheet renamed to 'Plan2'",
        "sheet:Plan2": "sheet renamed from 'Plan'",
    }


def test_unrelated_add_and_remove_keeps_neutral_wording(tmp_path):
    """Two unrelated sheets still get compared, but are not called a rename."""
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    _sheet_with_values(baseline, "Plan", {"A1": "Budget", "B1": 1000})
    _sheet_with_values(candidate, "Notes", {"A1": "totally", "B1": "different"})

    changes = diff(baseline, candidate)
    details = {c.location: c.detail for c in changes if c.location.startswith("sheet:")}
    assert details == {
        "sheet:Plan": "sheet removed; contents compared against 'Notes'",
        "sheet:Notes": "sheet added; contents compared against 'Plan'",
    }
    assert any(c.location == "Plan!A1" for c in changes)


def test_genuinely_added_sheet_is_not_paired(tmp_path):
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    _sheet_with_values(baseline, "Plan", {"A1": "Budget"})
    wb = Workbook()
    ws = wb.active
    ws.title = "Plan"
    ws["A1"] = "Budget"
    wb.create_sheet("Extra")["A1"] = "new"
    wb.save(candidate)

    changes = diff(baseline, candidate)
    assert [(c.location, c.detail) for c in changes] == [
        ("sheet:Extra", "sheet added")
    ]


def _sheet_with_rules(path, cf_formula, dv_formula):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws.conditional_formatting.add(
        "B2:B5 D2:D5",
        CellIsRule(operator="greaterThan", formula=[cf_formula],
                   fill=PatternFill("solid", fgColor="FFC7CE")),
    )
    dv = DataValidation(type="whole", operator="greaterThan", formula1=dv_formula)
    dv.add("B2:B5")
    dv.add("D2:D5")
    ws.add_data_validation(dv)
    wb.save(path)


def test_multi_range_rules_report_every_range(tmp_path):
    """A sqref like "B2:B5 D2:D5" must not collapse to its first range."""
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    _sheet_with_rules(baseline, "10", "0")
    _sheet_with_rules(candidate, "999", "500")

    locations = {(c.location, c.attribute) for c in diff(baseline, candidate)}
    assert ("S!B2:B5", ATTR_CONDITIONAL_FORMATTING) in locations
    assert ("S!D2:D5", ATTR_CONDITIONAL_FORMATTING) in locations
    assert ("S!B2:B5", ATTR_DATA_VALIDATION) in locations
    assert ("S!D2:D5", ATTR_DATA_VALIDATION) in locations

    # A policy that only opens up column B must not wave the D rule through.
    policy = parse_policy({
        "target": "excel",
        "allow": [{"selector": "S!B1:B100",
                   "attributes": ["conditional_formatting", "data_validation"]}],
    })
    result = check(baseline, candidate, policy)
    assert not result.passed
    assert {v.change.location for v in result.violations} == {"S!D2:D5"}


def test_result_json_roundtrip(fixtures):
    import json

    from templategate.reporters import render_json

    result = check(fixtures["excel_baseline"], fixtures["excel_bad"],
                   fixtures["excel_policy"])
    data = json.loads(render_json(result))
    assert data["tool"] == "templategate"
    assert data["passed"] is False
    assert data["summary"]["violations"] == len(result.violations)
