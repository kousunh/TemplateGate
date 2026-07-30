from docgate import check, diff, snapshot


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


def test_result_json_roundtrip(fixtures):
    import json

    from docgate.reporters import render_json

    result = check(fixtures["excel_baseline"], fixtures["excel_bad"],
                   fixtures["excel_policy"])
    data = json.loads(render_json(result))
    assert data["tool"] == "docgate"
    assert data["passed"] is False
    assert data["summary"]["violations"] == len(result.violations)
