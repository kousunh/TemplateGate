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


# --- default deny: parts nobody named still get compared ----------------

def _locations(changes, attribute):
    return {c.location for c in changes if c.attribute == attribute}


def test_unknown_part_is_tracked_by_the_catch_all(fixtures, tmp_path):
    """A part in no named category is still hashed, not silently dropped."""
    from generate import rewrite_zip

    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    rewrite_zip(fixtures["excel_baseline"], baseline,
                add={"xl/tables/table1.xml": b"<table ref='A1:B3'/>"})
    rewrite_zip(baseline, candidate, drop=("xl/tables/table1.xml",))

    changes = diff(baseline, candidate)
    assert _locations(changes, "parts") == {"package#parts:xl/tables/table1.xml"}
    assert not check(baseline, candidate, STRICT_EXCEL).passed


def test_external_link_part_content_is_compared(fixtures, tmp_path):
    """openpyxl cannot model external links; a repriced one must not slip by."""
    from generate import rewrite_zip

    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    part = "xl/externalLinks/externalLink1.xml"
    rewrite_zip(fixtures["excel_baseline"], baseline,
                add={part: b"<externalLink><v>9.99</v></externalLink>"})
    rewrite_zip(baseline, candidate,
                add={part: b"<externalLink><v>0.01</v></externalLink>"})
    assert _locations(diff(baseline, candidate), "parts") == {f"package#parts:{part}"}


def test_worksheet_extensions_are_compared(fixtures, tmp_path):
    """x14 dropdowns live in <extLst>, which openpyxl drops on every save."""
    from generate import rewrite_zip

    x14 = (
        '<extLst><ext uri="{CCE6A557}"><x14:dataValidations '
        'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
        "<x14:dataValidation type=\"list\"/></x14:dataValidations></ext></extLst>"
    )
    with zipfile.ZipFile(fixtures["excel_baseline"]) as zf:
        sheet = zf.read("xl/worksheets/sheet1.xml").decode()
    baseline = tmp_path / "base.xlsx"
    rewrite_zip(fixtures["excel_baseline"], baseline, add={
        "xl/worksheets/sheet1.xml":
            sheet.replace("</worksheet>", x14 + "</worksheet>").encode()
    })
    # Re-saving through openpyxl silently discards the extension block.
    from openpyxl import load_workbook

    candidate = tmp_path / "cand.xlsx"
    load_workbook(baseline).save(candidate)

    assert "package#parts:xl/worksheets/sheet1.xml#extLst" in _locations(
        diff(baseline, candidate), "parts")


def test_worksheet_body_is_not_compared_as_a_blob(fixtures, tmp_path):
    """Only the extensions of a worksheet part are hashed, never the cells."""
    package = take_package_snapshot(fixtures["excel_baseline"])
    assert not any(name.startswith("xl/worksheets/") and "#" not in name
                   for name in package["parts"])


def test_external_relationship_targets_are_compared(fixtures, tmp_path):
    """A hyperlink retargeted to another host changes no part content at all."""
    from generate import rewrite_zip

    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships"><Relationship Id="rId1" Type="http://x/hyperlink" '
        'Target="{url}" TargetMode="External"/></Relationships>'
    )
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    part = "xl/worksheets/_rels/sheet1.xml.rels"
    rewrite_zip(fixtures["excel_baseline"], baseline,
                add={part: rels.format(url="https://intranet.example.com/policy").encode()})
    rewrite_zip(baseline, candidate,
                add={part: rels.format(url="https://evil.example.net/phish").encode()})

    links = _locations(diff(baseline, candidate), "links")
    assert links == {
        "package#links:https://intranet.example.com/policy",
        "package#links:https://evil.example.net/phish",
    }
    assert not check(baseline, candidate, STRICT_EXCEL).passed


def test_relationship_ids_are_not_part_of_the_link_identity(fixtures, tmp_path):
    """Ids are renumbered on every rewrite; only targets are compared."""
    from generate import rewrite_zip

    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships"><Relationship Id="{rid}" Type="http://x/hyperlink" '
        'Target="https://example.com/a" TargetMode="External"/></Relationships>'
    )
    baseline = tmp_path / "base.xlsx"
    candidate = tmp_path / "cand.xlsx"
    part = "xl/worksheets/_rels/sheet1.xml.rels"
    rewrite_zip(fixtures["excel_baseline"], baseline,
                add={part: rels.format(rid="rId1").encode()})
    rewrite_zip(baseline, candidate, add={part: rels.format(rid="rId7").encode()})
    assert diff(baseline, candidate) == []


def test_custom_document_properties_are_tracked(fixtures, tmp_path):
    """docProps/custom.xml holds what DOCPROPERTY fields display."""
    from generate import rewrite_zip

    baseline = tmp_path / "base.docx"
    candidate = tmp_path / "cand.docx"
    rewrite_zip(fixtures["word_baseline"], baseline,
                add={"docProps/custom.xml": b"<Properties><p>APPROVED</p></Properties>"})
    rewrite_zip(baseline, candidate,
                add={"docProps/custom.xml": b"<Properties><p>DRAFT</p></Properties>"})
    assert _locations(diff(baseline, candidate), "parts") == {
        "package#parts:docProps/custom.xml"
    }


def test_document_timestamps_are_not_tracked(fixtures, tmp_path):
    """core.xml and app.xml change on every save and mean nothing."""
    from generate import rewrite_zip

    baseline = tmp_path / "base.docx"
    candidate = tmp_path / "cand.docx"
    rewrite_zip(fixtures["word_baseline"], baseline, add={})
    rewrite_zip(baseline, candidate, add={
        "docProps/core.xml": b"<coreProperties><modified>2031</modified></coreProperties>",
        "docProps/app.xml": b"<Properties><Company>x</Company></Properties>",
    })
    assert diff(baseline, candidate) == []


def test_word_revision_ids_are_normalized_away(fixtures, tmp_path):
    """Word stamps rsid attributes and proofing marks on every save."""
    from generate import rewrite_zip

    with zipfile.ZipFile(fixtures["word_baseline"]) as zf:
        styles = zf.read("word/styles.xml").decode()
    noisy = styles.replace("<w:style ", '<w:style w:rsidR="00AB12CD" ', 1)
    noisy = noisy.replace("</w:styles>", "<w:proofErr w:type='spellStart'/></w:styles>")

    baseline = tmp_path / "base.docx"
    candidate = tmp_path / "cand.docx"
    rewrite_zip(fixtures["word_baseline"], baseline, add={})
    rewrite_zip(baseline, candidate, add={"word/styles.xml": noisy.encode()})
    assert diff(baseline, candidate) == []


def test_word_style_definition_change_is_still_caught(fixtures, tmp_path):
    """Normalization must not blunt a real edit to a style definition."""
    from generate import rewrite_zip

    with zipfile.ZipFile(fixtures["word_baseline"]) as zf:
        styles = zf.read("word/styles.xml").decode()
    tampered = styles.replace("</w:styles>",
                              '<w:style w:styleId="Evil"><w:name w:val="Evil"/>'
                              "</w:style></w:styles>")
    baseline = tmp_path / "base.docx"
    candidate = tmp_path / "cand.docx"
    rewrite_zip(fixtures["word_baseline"], baseline, add={})
    rewrite_zip(baseline, candidate, add={"word/styles.xml": tampered.encode()})
    assert _locations(diff(baseline, candidate), "parts") == {
        "package#parts:word/styles.xml"
    }


def test_package_snapshot_is_json_serializable(fixtures):
    snap = snapshot(fixtures["excel_package_baseline"])
    assert set(snap["package"]) == set(CATEGORIES)
    json.loads(json.dumps(snap, default=str))
