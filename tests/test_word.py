from templategate import check, diff, snapshot
from templategate.core.model import ATTR_FORMAT, ATTR_TEXT
from templategate.core.policy import parse_policy


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


def test_run_formatting_is_captured(fixtures):
    """4pt white text leaves the paragraph text and style untouched."""
    snap = snapshot(fixtures["word_baseline"])
    assert "format" in snap["paragraphs"][0]
    assert snap["tables"][0]["cell_formats"]


def test_invisible_text_tampering_is_detected(fixtures):
    changes = diff(fixtures["word_baseline"], fixtures["word_reformatted"])
    assert [(c.location, c.attribute) for c in changes] == [("p4", ATTR_FORMAT)]

    # The text itself is unchanged, so a text-only allow rule must not cover it.
    policy = parse_policy({
        "target": "word",
        "allow": [{"selector": "body", "attributes": ["text"]}],
        "protect": [{"selector": "*", "attributes": ["style", "format"]}],
    })
    result = check(fixtures["word_baseline"], fixtures["word_reformatted"], policy)
    assert not result.passed
    assert result.violations[0].change.attribute == ATTR_FORMAT


def test_allowed_text_edit_does_not_report_a_format_change(fixtures):
    """Rewriting a paragraph must not fire a spurious format violation."""
    changes = diff(fixtures["word_baseline"], fixtures["word_good"])
    assert [(c.location, c.attribute) for c in changes] == [("p4", ATTR_TEXT)]


def test_paragraph_insertion_cascades_by_default(fixtures):
    changes = diff(fixtures["word_baseline"], fixtures["word_inserted"])
    assert len(changes) == 3
    assert {c.location for c in changes} == {"p3", "p4", "p5"}


def test_page_extension_aligns_paragraphs(fixtures):
    changes = diff(fixtures["word_baseline"], fixtures["word_inserted"], align=True)
    assert len(changes) == 1
    assert changes[0].location == "p3"
    assert changes[0].detail == "paragraph added"

    policy = parse_policy({
        "target": "word",
        "mode": "page_extension",
        "allow": [{"selector": "p3", "attributes": ["text"]}],
    })
    result = check(fixtures["word_baseline"], fixtures["word_inserted"], policy)
    assert result.passed, [v.message for v in result.violations]
