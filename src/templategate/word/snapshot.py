"""Word (.docx) snapshot extraction (read-only).

Content comes from the XML (see content.py), not from python-docx's object
graph, which silently omits whole classes of container.  python-docx is still
used to open the package, resolve style names and reach the relationships.
"""

from __future__ import annotations

import hashlib
from pathlib import Path

from docx import Document

from ..core.package import list_part_names, take_package_snapshot
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
        "images": sorted(images),
        "package": take_package_snapshot(path),
        "part_names": list_part_names(path),
    }
