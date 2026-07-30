"""Word (.docx) snapshot extraction (read-only)."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

from docx import Document


def _s(value) -> str:
    """Stringify a font property so signatures stay comparable and sortable."""
    return "" if value is None else str(value)


def _run_format(run) -> tuple[str, ...]:
    """Normalize one run's direct formatting into a comparable tuple."""
    f = run.font
    color = f.color
    return (
        _s(f.name),
        _s(int(f.size) if f.size is not None else None),
        _s(f.bold), _s(f.italic), _s(f.underline), _s(f.strike),
        _s(f.subscript), _s(f.superscript), _s(f.all_caps), _s(f.small_caps),
        _s(getattr(color, "rgb", None)),
        _s(getattr(color, "theme_color", None)),
        _s(f.highlight_color),
    )


def _format_key(paragraph) -> list[tuple[str, ...]]:
    """The set of distinct run formats used in a paragraph.

    Deduplicated and sorted on purpose: re-flowing the same text across a
    different number of runs is not a formatting change, while making any of
    the text bold, white, 4pt or highlighted is.  Keeping it independent of
    run boundaries means an allowed text edit does not also report a
    spurious format change.
    """
    return sorted({_run_format(r) for r in paragraph.runs})


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


def take_snapshot(path: str | Path) -> dict:
    path = Path(path)
    doc = Document(str(path))

    paragraphs = [
        {
            "text": p.text,
            "style": p.style.name if p.style else None,
            "format": _format_key(p),
        }
        for p in doc.paragraphs
    ]

    tables = []
    for table in doc.tables:
        rows = [[cell.text for cell in row.cells] for row in table.rows]
        cell_formats = [
            [sorted({fmt for p in cell.paragraphs for fmt in _format_key(p)})
             for cell in row.cells]
            for row in table.rows
        ]
        tables.append({
            "style": table.style.name if table.style else None,
            "rows": rows,
            "cell_formats": cell_formats,
        })

    sections = [_section(sec) for sec in doc.sections]
    header_footer = [_header_footer_text(sec) for sec in doc.sections]

    images = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if name.startswith("word/media/"):
                images.append(hashlib.sha256(zf.read(name)).hexdigest())

    return {
        "target": "word",
        "format": "docx",
        "paragraphs": paragraphs,
        "tables": tables,
        "sections": sections,
        "header_footer": header_footer,
        "images": sorted(images),
    }
