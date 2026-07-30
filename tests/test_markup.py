"""The backstop for markup nothing else models.

Every feature captured by name is a feature someone thought of.  These tests
cover the ones nobody had: a character style, a right-to-left mark, character
scaling, a raised baseline, a footnote reference, a comment anchor, a legacy
form field.  None of them needed a new rule — they are caught because what is
left after removing everything modelled is compared too.

The same tests pin the other half of the bargain: an allowed text edit must
not trip the backstop, or the whole thing is unusable.
"""

import zipfile

import pytest
from docx import Document

from templategate import check, diff, snapshot
from templategate.core.policy import parse_policy

STRICT = parse_policy({
    "target": "word",
    "protect": [{"selector": "*", "attributes": ["*"]}],
})
BODY_TEXT_ONLY = parse_policy({
    "target": "word",
    "allow": [{"selector": "body", "attributes": ["text"]}],
    "protect": [{"selector": "*", "attributes": ["style", "format", "table"]}],
})


def _build(tmp_path, name, body_xml):
    from generate import rewrite_zip

    plain = tmp_path / f"{name}_plain.docx"
    document = Document()
    document.add_paragraph("Intro paragraph.")
    document.save(str(plain))
    with zipfile.ZipFile(plain) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    built = tmp_path / f"{name}.docx"
    rewrite_zip(plain, built, add={
        "word/document.xml":
            xml.replace("</w:body>", body_xml + "</w:body>").encode("utf-8")})
    return built


def _pair(tmp_path, before, after):
    return _build(tmp_path, "base", before), _build(tmp_path, "cand", after)


def _attributes(changes):
    return {c.attribute for c in changes}


def _run(properties, text="Reference INV-2026-00417"):
    return f"<w:p><w:r>{properties}<w:t>{text}</w:t></w:r></w:p>"


@pytest.mark.parametrize("name,before,after", [
    ("character style",
     "<w:rPr><w:rStyle w:val=\"Strong\"/></w:rPr>", "<w:rPr></w:rPr>"),
    ("right to left",
     "", "<w:rPr><w:rtl/></w:rPr>"),
    ("character scaling",
     "", "<w:rPr><w:w w:val=\"10\"/></w:rPr>"),
    ("raised baseline",
     "", "<w:rPr><w:position w:val=\"-60\"/></w:rPr>"),
])
def test_unmodelled_run_property_is_caught(tmp_path, name, before, after):
    baseline, candidate = _pair(tmp_path, _run(before), _run(after))
    changes = diff(baseline, candidate)
    assert "markup" in _attributes(changes), name
    assert not check(baseline, candidate, STRICT).passed


def test_footnote_reference_removed_from_the_body_is_caught(tmp_path):
    """The footnote text stays in the package; the page loses the mark."""
    reference = ('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                 '<w:footnoteReference w:id="1"/></w:r>')
    body = "<w:p><w:r><w:t>Record revenue.</w:t></w:r>" + reference + "</w:p>"
    plain = "<w:p><w:r><w:t>Record revenue.</w:t></w:r></w:p>"
    baseline, candidate = _pair(tmp_path, body, plain)
    assert "markup" in _attributes(diff(baseline, candidate))
    assert not check(baseline, candidate, BODY_TEXT_ONLY).passed


def test_comment_anchor_removed_from_the_body_is_caught(tmp_path):
    anchored = ('<w:p><w:commentRangeStart w:id="1"/>'
                "<w:r><w:t>May subcontract.</w:t></w:r>"
                '<w:commentRangeEnd w:id="1"/>'
                '<w:r><w:commentReference w:id="1"/></w:r></w:p>')
    bare = "<w:p><w:r><w:t>May subcontract.</w:t></w:r></w:p>"
    baseline, candidate = _pair(tmp_path, anchored, bare)
    assert "markup" in _attributes(diff(baseline, candidate))


def test_legacy_form_field_disabled_is_caught(tmp_path):
    def form(enabled, maxlen):
        disabled = "" if enabled else '<w:enabled w:val="0"/>'
        return ('<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData>'
                f'<w:name w:val="ApplicantName"/>{disabled}'
                f'<w:textInput><w:maxLength w:val="{maxlen}"/></w:textInput>'
                "</w:ffData></w:fldChar></w:r>"
                "<w:r><w:t>(enter name)</w:t></w:r></w:p>")

    baseline, candidate = _pair(tmp_path, form(True, "40"), form(False, "2"))
    assert "markup" in _attributes(diff(baseline, candidate))


# --- and the other half: no false positives -----------------------------

def test_rewriting_body_text_reports_text_only(tmp_path):
    """The backstop must not fire on the edit a policy exists to allow."""
    baseline, candidate = _pair(
        tmp_path,
        "<w:p><w:r><w:t>Original sentence here.</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>Rewritten sentence here.</w:t></w:r></w:p>")
    assert [(c.location, c.attribute) for c in diff(baseline, candidate)] == [
        ("p2", "text")
    ]
    assert check(baseline, candidate, BODY_TEXT_ONLY).passed


def test_resplitting_the_same_text_across_runs_is_not_a_change(tmp_path):
    """Editors re-flow runs constantly; that is not an edit."""
    one_run = "<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>"
    two_runs = ("<w:p><w:r><w:t>Hello </w:t></w:r>"
                "<w:r><w:t>world</w:t></w:r></w:p>")
    baseline, candidate = _pair(tmp_path, one_run, two_runs)
    assert diff(baseline, candidate) == []


def test_replacing_text_across_several_runs_reports_text_only(tmp_path):
    """Filling placeholders run by run is the classic agent edit."""
    before = ("<w:p><w:r><w:t>Dear </w:t></w:r>"
              "<w:r><w:t>{{NAME}}</w:t></w:r>"
              "<w:r><w:t>,</w:t></w:r></w:p>")
    after = ("<w:p><w:r><w:t>Dear </w:t></w:r>"
             "<w:r><w:t>Ms Ito</w:t></w:r>"
             "<w:r><w:t>,</w:t></w:r></w:p>")
    baseline, candidate = _pair(tmp_path, before, after)
    assert _attributes(diff(baseline, candidate)) == {"text"}


def test_proofing_metadata_does_not_make_a_text_edit_a_failure(tmp_path):
    """Real Word stamps w:lang on every run a person touches.

    It changes nothing a reader can see, so an approved text edit made in Word
    must still report one text change and nothing else.
    """
    before = "<w:p><w:r><w:t>Original sentence here.</w:t></w:r></w:p>"
    after = ('<w:p><w:r><w:rPr><w:lang w:val="en-US"/><w:noProof/>'
             '<w:snapToGrid w:val="0"/></w:rPr>'
             "<w:t>Rewritten sentence here.</w:t></w:r></w:p>")
    baseline, candidate = _pair(tmp_path, before, after)
    assert [(c.location, c.attribute) for c in diff(baseline, candidate)] == [
        ("p2", "text")
    ]
    assert check(baseline, candidate, BODY_TEXT_ONLY).passed


def test_proofing_metadata_alone_is_not_a_change(tmp_path):
    before = "<w:p><w:r><w:t>Clause.</w:t></w:r></w:p>"
    after = ('<w:p><w:r><w:rPr><w:lang w:val="ja-JP"/></w:rPr>'
             "<w:t>Clause.</w:t></w:r></w:p>")
    baseline, candidate = _pair(tmp_path, before, after)
    assert diff(baseline, candidate) == []


@pytest.mark.parametrize("properties,name", [
    ("<w:rPr><w:specVanish/></w:rPr>", "specVanish"),
    ('<w:rPr><w:kern w:val="2"/></w:rPr>', "kern"),
])
def test_rendering_properties_are_still_markup(tmp_path, properties, name):
    """Ignoring proofing metadata must not quietly ignore its neighbours."""
    baseline, candidate = _pair(
        tmp_path,
        "<w:p><w:r><w:t>Clause.</w:t></w:r></w:p>",
        f"<w:p><w:r>{properties}<w:t>Clause.</w:t></w:r></w:p>")
    changes = diff(baseline, candidate)
    assert "markup" in _attributes(changes), name
    assert not check(baseline, candidate, BODY_TEXT_ONLY).passed


def test_revision_ids_are_not_markup(tmp_path):
    """Word stamps rsid attributes on every save; they mean nothing."""
    before = "<w:p><w:r><w:t>Clause.</w:t></w:r></w:p>"
    after = ('<w:p w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD">'
             '<w:r w:rsidRPr="00FF0011"><w:t>Clause.</w:t></w:r></w:p>')
    baseline, candidate = _pair(tmp_path, before, after)
    assert diff(baseline, candidate) == []


def test_tracked_insertion_is_not_reported_as_markup(tmp_path):
    """It is already a text change; saying so twice helps nobody."""
    plain = "<w:p><w:r><w:t>The Supplier is liable.</w:t></w:r></w:p>"
    inserted = ('<w:p><w:r><w:t>The Supplier is liable.</w:t></w:r>'
                '<w:ins w:id="9" w:author="a" w:date="2026-01-01T00:00:00Z">'
                "<w:r><w:t> Except where not.</w:t></w:r></w:ins></w:p>")
    baseline, candidate = _pair(tmp_path, plain, inserted)
    assert "markup" not in _attributes(diff(baseline, candidate))


def test_swapping_a_picture_reports_images_only(fixtures):
    """Relationship ids and source filenames churn; content is what matters."""
    changes = diff(fixtures["word_image_baseline"], fixtures["word_image_swapped"])
    assert _attributes(changes) == {"images"}


def test_markup_is_captured_for_every_container(tmp_path):
    """Content controls, text boxes and table cells get the backstop too."""
    body = (
        '<w:sdt><w:sdtPr><w:tag w:val="a"/></w:sdtPr><w:sdtContent>'
        "<w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p>"
        "</w:tc></w:tr></w:tbl>"
    )
    document = _build(tmp_path, "doc", body)
    snap = snapshot(document)
    assert "markup" in snap["paragraphs"][0]
    assert "markup" in snap["blocks"]["sdt1"]
    assert snap["tables"][0]["cell_markup"]


def test_markup_in_a_table_cell_is_caught(tmp_path):
    def cell(properties):
        return ("<w:tbl><w:tr><w:tc><w:p>"
                f"<w:r>{properties}<w:t>300</w:t></w:r>"
                "</w:p></w:tc></w:tr></w:tbl>")

    baseline, candidate = _pair(tmp_path, cell(""),
                                cell("<w:rPr><w:rtl/></w:rPr>"))
    changes = [c for c in diff(baseline, candidate) if c.attribute == "markup"]
    assert [c.location for c in changes] == ["table1!r1c1"]
