"""OOXML package-part detection.

Charts, pivot tables, shapes, comments, embedded objects and VBA are invisible
to openpyxl and python-docx — they are dropped on save without a word — so the
gate reads them straight from the zip.  These tests pin both halves of that:
the damage is caught, and a legitimate re-save stays silent.
"""

import json
import zipfile

import pytest

from templategate import check, diff, snapshot
from templategate.core.package import CATEGORIES, take_package_snapshot
from templategate.core.policy import parse_policy

STRICT_EXCEL = parse_policy({
    "target": "excel",
    "protect": [{"selector": "*", "attributes": ["*"]}],
})
STRICT_WORD = parse_policy({
    "target": "word",
    "protect": [{"selector": "*", "attributes": ["*"]}],
})

EXCEL_LOCATIONS = {
    "vba": {"vba"},
    "charts": {"package#charts:xl/charts/chart1.xml"},
    "pivot_tables": {
        "package#pivot_tables:xl/pivotTables/pivotTable1.xml",
        "package#pivot_tables:xl/pivotCache/pivotCacheDefinition1.xml",
    },
    "comments": {
        "package#comments:xl/comments/comment1.xml",
        "package#comments:xl/threadedComments/threadedComment1.xml",
    },
    "embedded": {"package#embedded:xl/embeddings/oleObject1.bin"},
    "custom_xml": {"package#custom_xml:customXml/item1.xml"},
    "drawings": {"package#drawings:xl/drawings/drawing1.xml"},
}


def _locations_by_attribute(changes):
    found: dict[str, set[str]] = {}
    for change in changes:
        found.setdefault(change.attribute, set()).add(change.location)
    return found


@pytest.mark.parametrize("category", sorted(EXCEL_LOCATIONS))
def test_excel_part_removal_is_detected(fixtures, category):
    """Every category survives the trip from 'present' to 'gone'."""
    found = _locations_by_attribute(
        diff(fixtures["excel_package_baseline"], fixtures["excel_package_stripped"])
    )
    assert found[category] == EXCEL_LOCATIONS[category]


@pytest.mark.parametrize("category", sorted(EXCEL_LOCATIONS))
def test_excel_part_addition_is_detected(fixtures, category):
    found = _locations_by_attribute(
        diff(fixtures["excel_package_stripped"], fixtures["excel_package_baseline"])
    )
    assert found[category] == EXCEL_LOCATIONS[category]


@pytest.mark.parametrize("category", sorted(EXCEL_LOCATIONS))
def test_excel_part_modification_is_detected(fixtures, category):
    found = _locations_by_attribute(
        diff(fixtures["excel_package_baseline"], fixtures["excel_package_modified"])
    )
    assert found[category] == EXCEL_LOCATIONS[category]


def test_removal_addition_and_modification_are_distinguishable(fixtures):
    removed = diff(fixtures["excel_package_baseline"], fixtures["excel_package_stripped"])
    modified = diff(fixtures["excel_package_baseline"], fixtures["excel_package_modified"])
    by_location = {c.location: c for c in removed}
    chart = by_location["package#charts:xl/charts/chart1.xml"]
    assert chart.new is None and chart.old is not None
    assert "removed" in chart.detail

    chart = {c.location: c for c in modified}["package#charts:xl/charts/chart1.xml"]
    assert chart.old is not None and chart.new is not None
    assert "modified" in chart.detail


def test_stripped_package_fails_a_strict_policy(fixtures):
    result = check(fixtures["excel_package_baseline"],
                   fixtures["excel_package_stripped"], STRICT_EXCEL)
    assert not result.passed
    assert {v.change.attribute for v in result.violations} >= set(EXCEL_LOCATIONS)


@pytest.mark.parametrize("category", sorted(set(CATEGORIES) - {"vba"}))
def test_structural_ignore_opts_out_of_one_category(fixtures, category):
    """Each category has a structural key that silences exactly itself."""
    policy = parse_policy({
        "target": "excel",
        "protect": [{"selector": "*", "attributes": ["*"]}],
        "structural": {category: "ignore"},
    })
    result = check(fixtures["excel_package_baseline"],
                   fixtures["excel_package_stripped"], policy)
    attributes = {v.change.attribute for v in result.violations}
    assert category not in attributes
    assert attributes  # the other categories still fail


def test_word_part_removal_is_detected(fixtures):
    found = _locations_by_attribute(
        diff(fixtures["word_package_baseline"], fixtures["word_package_stripped"])
    )
    assert found["vba"] == {"vba"}
    assert found["charts"] == {"package#charts:word/charts/chart1.xml"}
    assert found["comments"] == {"package#comments:word/comments.xml"}
    assert found["embedded"] == {"package#embedded:word/embeddings/oleObject1.bin"}
    result = check(fixtures["word_package_baseline"],
                   fixtures["word_package_stripped"], STRICT_WORD)
    assert not result.passed


def test_word_part_modification_is_detected(fixtures):
    found = _locations_by_attribute(
        diff(fixtures["word_package_baseline"], fixtures["word_package_modified"])
    )
    assert found["custom_xml"] == {"package#custom_xml:customXml/item1.xml"}
    assert found["comments"] == {"package#comments:word/comments.xml"}


def test_vba_keeps_its_standalone_location(fixtures):
    """Policies written against the old "vba" selector must keep working."""
    policy = parse_policy({
        "target": "excel",
        "allow": [{"selector": "package#*", "attributes": ["*"]}],
        "protect": [{"selector": "vba", "attributes": ["vba"]}],
    })
    result = check(fixtures["excel_package_baseline"],
                   fixtures["excel_package_stripped"], policy)
    vba = [v for v in result.violations if v.change.attribute == "vba"]
    assert [v.rule for v in vba] == ["protected"]
    assert vba[0].change.location == "vba"


# --- no false positives -------------------------------------------------

def test_legitimate_roundtrip_reports_only_the_edited_cell(fixtures):
    """openpyxl load + one cell edit + save must not trip the package layer.

    openpyxl rewrites drawing XML on every save (it injects default children
    such as <a:ln>), so comparing those parts byte-for-byte would report a
    violation for an edit the policy explicitly allows.
    """
    changes = diff(fixtures["excel_native_rich"], fixtures["excel_native_edited"])
    assert [(c.location, c.attribute) for c in changes] == [("Data!B2", "value")]


def test_volatile_parts_are_ignored(fixtures):
    """Rewritten timestamps, calc chain, printer settings and orphan media."""
    assert diff(fixtures["excel_package_baseline"],
                fixtures["excel_package_volatile"]) == []


def test_scanner_skips_excluded_parts(tmp_path):
    """Content types, relationship files and media never enter the inventory."""
    archive = tmp_path / "pkg.zip"
    with zipfile.ZipFile(archive, "w") as zf:
        zf.writestr("[Content_Types].xml", b"<Types/>")
        zf.writestr("_rels/.rels", b"<Relationships/>")
        zf.writestr("xl/charts/_rels/chart1.xml.rels", b"<Relationships/>")
        zf.writestr("docProps/core.xml", b"<coreProperties/>")
        zf.writestr("xl/calcChain.xml", b"<calcChain/>")
        zf.writestr("xl/printerSettings/printerSettings1.bin", b"blob")
        zf.writestr("xl/media/image1.png", b"pixels")
        zf.writestr("xl/charts/chart1.xml", b"<c:chartSpace/>")

    package = take_package_snapshot(archive)
    assert package["charts"] == {
        "xl/charts/chart1.xml": package["charts"]["xl/charts/chart1.xml"]
    }
    assert list(package["charts"]) == ["xl/charts/chart1.xml"]
    assert all(not parts for name, parts in package.items() if name != "charts")


def _drawing(shapes: str) -> bytes:
    return (
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006'
        '/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml'
        f'/2006/main">{shapes}</xdr:wsDr>'
    ).encode()


def _package_with_drawing(path, body: str):
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/drawings/drawing1.xml", _drawing(body))
    return take_package_snapshot(path)["drawings"]


def test_drawing_digest_ignores_injected_presentation_defaults(tmp_path):
    """The exact churn openpyxl produces must not read as a change."""
    plain = _package_with_drawing(
        tmp_path / "a.xlsx",
        "<xdr:twoCellAnchor><xdr:sp><xdr:txBody><a:p><a:r><a:t>SEAL</a:t>"
        "</a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor>",
    )
    with_defaults = _package_with_drawing(
        tmp_path / "b.xlsx",
        "<xdr:twoCellAnchor><xdr:sp><xdr:spPr><a:ln><a:prstDash val=\"solid\"/>"
        "</a:ln></xdr:spPr><xdr:txBody><a:p><a:r><a:t>SEAL</a:t>"
        "</a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor>",
    )
    assert plain == with_defaults


def test_drawing_digest_still_catches_a_removed_shape(tmp_path):
    before = _package_with_drawing(
        tmp_path / "a.xlsx",
        "<xdr:twoCellAnchor><xdr:sp><xdr:txBody><a:p><a:r><a:t>SEAL</a:t>"
        "</a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor>",
    )
    after = _package_with_drawing(tmp_path / "b.xlsx", "")
    assert before != after


def test_drawing_digest_catches_edited_shape_text(tmp_path):
    before = _package_with_drawing(
        tmp_path / "a.xlsx",
        "<xdr:sp><xdr:txBody><a:p><a:r><a:t>APPROVED</a:t></a:r></a:p>"
        "</xdr:txBody></xdr:sp>",
    )
    after = _package_with_drawing(
        tmp_path / "b.xlsx",
        "<xdr:sp><xdr:txBody><a:p><a:r><a:t>REJECTED</a:t></a:r></a:p>"
        "</xdr:txBody></xdr:sp>",
    )
    assert before != after


def test_picture_only_drawing_is_not_tracked(tmp_path):
    """Pictures are the images attribute's job; tracking them here would
    report every added or deleted image twice."""
    drawings = _package_with_drawing(
        tmp_path / "a.xlsx",
        "<xdr:oneCellAnchor><xdr:pic><xdr:blipFill/></xdr:pic></xdr:oneCellAnchor>",
    )
    assert drawings == {}


def test_unparseable_drawing_falls_back_to_hashing_bytes(tmp_path):
    """A part we cannot parse must still be compared, not silently skipped."""
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"
    with zipfile.ZipFile(a, "w") as zf:
        zf.writestr("xl/drawings/drawing1.xml", b"<not valid xml")
    with zipfile.ZipFile(b, "w") as zf:
        zf.writestr("xl/drawings/drawing1.xml", b"<also not valid xml")
    assert take_package_snapshot(a)["drawings"]
    assert take_package_snapshot(a) != take_package_snapshot(b)


def test_package_snapshot_is_json_serializable(fixtures):
    snap = snapshot(fixtures["excel_package_baseline"])
    assert set(snap["package"]) == set(CATEGORIES)
    json.loads(json.dumps(snap, default=str))
