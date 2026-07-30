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


def test_array_formula_is_protected_as_a_formula(tmp_path):
    """An array formula is not a str, and used to be filed under `value`.

    A policy that allows value edits then let =A2:A4*C2:C4 become =A2:A4*0.
    """
    from openpyxl.worksheet.formula import ArrayFormula

    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    for path, text in ((baseline, "=A2:A4*C2:C4"), (candidate, "=A2:A4*0")):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A2"], ws["C2"] = 10, 50
        ws["B2"] = ArrayFormula("B2:B4", text)
        wb.save(path)

    changes = diff(baseline, candidate)
    assert [(c.location, c.attribute) for c in changes] == [("Sheet1!B2", ATTR_FORMULA)]

    policy = parse_policy({
        "target": "excel",
        "allow": [{"selector": "*", "attributes": ["value"]}],
        "protect": [{"selector": "*", "attributes": ["formula"]}],
    })
    result = check(baseline, candidate, policy)
    assert not result.passed
    assert result.violations[0].rule == "protected"


def test_array_formula_survives_an_untouched_comparison(tmp_path):
    """ArrayFormula defines no __eq__; comparing a file with itself must be quiet."""
    from openpyxl.worksheet.formula import ArrayFormula

    path = tmp_path / "base.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 1
    ws["D1"] = ArrayFormula("D1:D3", "=A1:A3*2")
    wb.save(path)
    assert diff(path, path) == []


def test_dates_are_stored_as_text(tmp_path):
    """No datetime object may reach the snapshot; JSON has to survive it."""
    import datetime
    import json

    path = tmp_path / "base.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = datetime.datetime(2026, 7, 31, 12, 30)
    ws["A2"] = datetime.time(9, 15)
    wb.save(path)

    cells = snapshot(path)["sheets"]["Sheet1"]["cells"]
    assert cells["A1"]["value"] == "2026-07-31T12:30:00"
    assert cells["A2"]["value"] == "09:15:00"
    json.dumps(snapshot(path))  # no default= crutch needed


def _chartsheet_workbook(path, values):
    from openpyxl.chart import BarChart, Reference

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row, value in enumerate(values, start=1):
        ws.cell(row=row, column=1, value=value)
    sheet = wb.create_chartsheet("Chart1")
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=1, min_row=1, max_row=len(values)))
    sheet.add_chart(chart)
    wb.save(path)


def test_chartsheet_does_not_crash(tmp_path):
    """A chartsheet has no cell grid; asking it for rows used to raise."""
    baseline = tmp_path / "base.xlsx"
    _chartsheet_workbook(baseline, [5, 3, 8, 2])

    snap = snapshot(baseline)
    assert snap["sheets"]["Chart1"]["kind"] == "chartsheet"
    assert snap["sheets"]["Chart1"]["cells"] == {}
    assert snap["sheets"]["Data"]["kind"] == "worksheet"
    assert diff(baseline, baseline) == []


def test_chartsheet_cell_edits_on_the_data_sheet_still_report(tmp_path):
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    _chartsheet_workbook(baseline, [5, 3, 8, 2])
    _chartsheet_workbook(candidate, [5, 3, 8, 99])
    assert [(c.location, c.attribute) for c in diff(baseline, candidate)] == [
        ("Data!A4", ATTR_VALUE)
    ]


def test_worksheet_replaced_by_chartsheet_is_reported(tmp_path):
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    _sheet_with_values(baseline, "Chart1", {"A1": "was a worksheet"})
    _chartsheet_workbook(candidate, [1, 2])
    details = {c.detail for c in diff(baseline, candidate)}
    assert "sheet kind changed" in details


def test_result_json_roundtrip(fixtures):
    import json

    from templategate.reporters import render_json

    result = check(fixtures["excel_baseline"], fixtures["excel_bad"],
                   fixtures["excel_policy"])
    data = json.loads(render_json(result))
    assert data["tool"] == "templategate"
    assert data["passed"] is False
    assert data["summary"]["violations"] == len(result.violations)
