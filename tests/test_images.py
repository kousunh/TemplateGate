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


# --- relationship targets written the way real writers write them --------

def _rewrite_targets(source, destination, transform):
    """Restate every worksheet relationship target in another legal form."""
    import zipfile

    from generate import rewrite_zip

    patched = {}
    with zipfile.ZipFile(source) as zf:
        for name in zf.namelist():
            if name.startswith("xl/worksheets/_rels/"):
                patched[name] = transform(
                    zf.read(name).decode("utf-8")).encode("utf-8")
    assert patched, "no worksheet relationships to rewrite"
    rewrite_zip(source, destination, add=patched)
    return destination


def test_a_relative_relationship_target_resolves(fixtures, tmp_path):
    """Real Excel writes Target="../drawings/drawing1.xml", openpyxl writes
    "/xl/drawings/drawing1.xml", and both name the same part.

    Joining the relative form onto the worksheet folder without normalising
    it yields "xl/worksheets/../drawings/drawing1.xml", which is not a member
    of the zip — so the drawing is never found and every image silently falls
    back to its intrinsic size.
    """
    import zipfile

    from templategate.excel.snapshot import _relationship_targets, _sheet_parts

    relative = _rewrite_targets(
        fixtures["excel_image_baseline"], tmp_path / "relative.xlsx",
        lambda xml: xml.replace('Target="/xl/drawings/', 'Target="../drawings/')
                       .replace('Target="xl/drawings/', 'Target="../drawings/'))

    with zipfile.ZipFile(relative) as zf:
        members = set(zf.namelist())
        resolved = [target
                    for part in _sheet_parts(zf).values()
                    for target in _relationship_targets(zf, part).values()]
    assert resolved, "the worksheet declared no relationships"
    assert all(target in members for target in resolved), resolved


def test_restating_a_target_relatively_is_not_a_change(fixtures, tmp_path):
    """The same package, spelled the other way, must compare equal — this is
    the shape of the false positive: an Excel-authored baseline against an
    openpyxl-resaved candidate.
    """
    relative = _rewrite_targets(
        fixtures["excel_image_baseline"], tmp_path / "relative.xlsx",
        lambda xml: xml.replace('Target="/xl/drawings/', 'Target="../drawings/')
                       .replace('Target="xl/drawings/', 'Target="../drawings/'))

    assert diff(fixtures["excel_image_baseline"], relative) == []
