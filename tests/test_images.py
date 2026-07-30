"""Regression tests: embedded images must never be invisible to the gate.

openpyxl silently drops every image while reading a workbook when Pillow is
not installed, which used to make image add/remove/replace a silent PASS.
"""

from templategate import check, diff
from templategate.core.model import ATTR_IMAGES
from templategate.core.policy import parse_policy

STRICT_EXCEL = parse_policy({
    "target": "excel",
    "protect": [{"selector": "*", "attributes": ["*"]}],
    "structural": {"images": "strict"},
})
STRICT_WORD = parse_policy({
    "target": "word",
    "protect": [{"selector": "*", "attributes": ["*"]}],
    "structural": {"images": "strict"},
})


def test_openpyxl_can_actually_read_images(fixtures):
    """Guards the Pillow dependency: without it this list is silently empty."""
    from templategate import snapshot

    snap = snapshot(fixtures["excel_image_baseline"])
    images = snap["sheets"]["表紙"]["images"]
    assert len(images) == 1
    assert images[0]["sha256"]


def test_excel_image_replaced_is_detected(fixtures):
    changes = diff(fixtures["excel_image_baseline"], fixtures["excel_image_swapped"])
    assert [c.attribute for c in changes] == [ATTR_IMAGES, ATTR_IMAGES]
    result = check(fixtures["excel_image_baseline"], fixtures["excel_image_swapped"],
                   STRICT_EXCEL)
    assert not result.passed


def test_excel_image_removed_is_detected(fixtures):
    changes = diff(fixtures["excel_image_baseline"], fixtures["excel_image_removed"])
    assert [c.attribute for c in changes] == [ATTR_IMAGES]
    assert changes[0].new is None
    result = check(fixtures["excel_image_baseline"], fixtures["excel_image_removed"],
                   STRICT_EXCEL)
    assert not result.passed


def test_excel_image_added_is_detected(fixtures):
    changes = diff(fixtures["excel_image_removed"], fixtures["excel_image_baseline"])
    assert [c.attribute for c in changes] == [ATTR_IMAGES]
    assert changes[0].old is None


def test_excel_images_ignore_setting_still_works(fixtures):
    policy = parse_policy({
        "target": "excel",
        "protect": [{"selector": "*", "attributes": ["*"]}],
        "structural": {"images": "ignore"},
    })
    result = check(fixtures["excel_image_baseline"], fixtures["excel_image_swapped"],
                   policy)
    assert result.passed


def test_word_image_replaced_is_detected(fixtures):
    changes = diff(fixtures["word_image_baseline"], fixtures["word_image_swapped"])
    assert [c.attribute for c in changes] == [ATTR_IMAGES, ATTR_IMAGES]
    result = check(fixtures["word_image_baseline"], fixtures["word_image_swapped"],
                   STRICT_WORD)
    assert not result.passed


def test_word_image_removed_is_detected(fixtures):
    changes = diff(fixtures["word_image_baseline"], fixtures["word_image_removed"])
    assert any(c.attribute == ATTR_IMAGES and c.new is None for c in changes)
