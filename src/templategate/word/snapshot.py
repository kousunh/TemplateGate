"""Word (.docx) snapshot extraction (read-only).

Content comes from the XML (see content.py), not from python-docx's object
graph, which silently omits whole classes of container.  python-docx is still
used to open the package, resolve style names and reach the relationships.
"""

from __future__ import annotations

import hashlib
from pathlib import Path

from docx import Document

from ..core.package import (
    canonical_xml_digest,
    list_part_names,
    take_package_snapshot,
)
from .content import walk_body


def _section(sec) -> dict:
    def emu(v):
        return int(v) if v is not None else None

    return {
        "orientation": str(sec.orientation),
        "page_width": emu(sec.page_width),
        "page_height": emu(sec.page_height),
        "margins": (emu(sec.left_margin), emu(sec.right_margin),
                    emu(sec.top_margin), emu(sec.bottom_margin)),
        "header_distance": emu(sec.header_distance),
        "footer_distance": emu(sec.footer_distance),
    }


def _header_footer_text(sec) -> dict:
    out = {}
    for name, part in (("header", sec.header), ("footer", sec.footer)):
        if part is None or part.is_linked_to_previous:
            continue
        text = "\n".join(p.text for p in part.paragraphs).strip()
        if text:
            out[name] = text
    return out


_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _displayed_images(doc) -> list[str]:
    """Hashes of the images the document actually shows.

    Comparing the media folder instead would miss a picture deleted from the
    page whose bytes are still sitting in the package — the file looks the
    same size, and the logo is gone.
    """
    parts = [doc.part]
    for section in doc.sections:
        for source in (section.header, section.footer,
                       section.even_page_header, section.even_page_footer,
                       section.first_page_header, section.first_page_footer):
            if source is not None and not source.is_linked_to_previous:
                parts.append(source.part)

    hashes: list[str] = []
    for part in parts:
        element = getattr(part, "element", None)
        if element is None:
            continue
        references = [
            reference.get(_R + "embed") or reference.get(_R + "link")
            for reference in element.iter(_BLIP)
        ] + [
            reference.get(_R + "id") for reference in element.iter(_IMAGEDATA)
        ]
        for rel_id in references:
            if not rel_id or rel_id not in part.rels:
                continue
            relationship = part.rels[rel_id]
            if relationship.is_external:
                hashes.append(f"external:{relationship.target_ref}")
                continue
            try:
                hashes.append(hashlib.sha256(relationship.target_part.blob).hexdigest())
            except Exception:
                continue
    return sorted(hashes)


def _header_footer_parts(doc) -> dict[str, str]:
    """Header and footer parts keyed by the role they fill.

    Word writes header1..3 and footer1..3, then discards the unused ones and
    renumbers what is left.  Old header2 becomes new header1 with identical
    bytes, so comparing by filename reports one modified and two removed for
    a document nobody touched.  What a reader actually cares about is which
    header is on the default page, the even page and the first page.
    """
    found: dict[str, str] = {}
    # A header only exists for a reader if the document is set to show it.
    # Word keeps even-page and first-page parts around while those options are
    # off and discards them on the next save, which is not a deletion.
    even_shown = bool(getattr(doc.settings, "odd_and_even_pages_header_footer",
                              False))
    for index, section in enumerate(doc.sections, start=1):
        roles = ("default",
                 "even" if even_shown else None,
                 "first" if section.different_first_page_header_footer else None)
        sources = (
            ("header", (section.header, section.even_page_header,
                        section.first_page_header)),
            ("footer", (section.footer, section.even_page_footer,
                        section.first_page_footer)),
        )
        for kind, parts in sources:
            for role, source in zip(roles, parts):
                if role is None or source is None or source.is_linked_to_previous:
                    continue
                part = getattr(source, "part", None)
                if part is None:
                    continue
                digest = canonical_xml_digest(part.blob)
                if digest is None:
                    digest = hashlib.sha256(part.blob).hexdigest()
                found[f"section{index}#{kind}:{role}"] = digest
    return found


def take_snapshot(path: str | Path) -> dict:
    path = Path(path)
    doc = Document(str(path))

    style_names = {}
    for style in doc.styles:
        style_id = getattr(style, "style_id", None)
        if style_id:
            style_names[style_id] = style.name

    paragraphs, tables, blocks = walk_body(doc.element.body, style_names)

    sections = [_section(sec) for sec in doc.sections]
    header_footer = [_header_footer_text(sec) for sec in doc.sections]
    images = _displayed_images(doc)

    return {
        "target": "word",
        "blocks": blocks,
        "format": "docx",
        "paragraphs": paragraphs,
        "tables": tables,
        "sections": sections,
        "header_footer": header_footer,
        "header_footer_parts": _header_footer_parts(doc),
        "images": sorted(images),
        "package": take_package_snapshot(path),
        "part_names": list_part_names(path),
    }
