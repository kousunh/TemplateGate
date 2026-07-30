from templategate import check, diff, snapshot


def test_snapshot_contents(fixtures):
    snap = snapshot(fixtures["word_baseline"])
    assert snap["target"] == "word"
    assert any("業務計画書" in p["text"] for p in snap["paragraphs"])
    assert snap["tables"][0]["rows"][1][1] == "2026年度〜2030年度"
    assert snap["sections"]


def test_identical_files_have_no_changes(fixtures):
    assert diff(fixtures["word_baseline"], fixtures["word_baseline"]) == []


def test_good_edit_passes(fixtures):
    result = check(fixtures["word_baseline"], fixtures["word_good"],
                   fixtures["word_policy"])
    assert result.passed, [v.message for v in result.violations]
    assert any(c.location == "p4" for c in result.allowed)


def test_bad_edit_fails_on_table_change(fixtures):
    result = check(fixtures["word_baseline"], fixtures["word_bad"],
                   fixtures["word_policy"])
    assert not result.passed
    locations = {v.change.location for v in result.violations}
    assert "table1!r2c2" in locations


def test_target_mismatch_rejected(fixtures):
    import pytest

    with pytest.raises(ValueError):
        check(fixtures["word_baseline"], fixtures["word_good"],
              fixtures["excel_policy"])
