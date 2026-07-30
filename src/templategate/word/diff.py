"""Diff two Word snapshots into Change records.

Paragraphs and tables are compared by position (1-based index).  Inserting a
paragraph in the middle therefore reports every following paragraph as
changed — deliberate for an acceptance gate: position shifts in a fixed
template are exactly what it must catch.  Policies for append-style edits can
allow a trailing paragraph range instead, or set ``mode: page_extension``,
which aligns paragraphs by content (``align=True`` here) so that an insertion
reports only the inserted paragraph.
"""

from __future__ import annotations

from collections import Counter
from difflib import SequenceMatcher

from ..core.model import (
    ATTR_BOOKMARK,
    ATTR_CONTENT_CONTROL,
    ATTR_FIELD,
    ATTR_FORMAT,
    ATTR_HEADER_FOOTER,
    ATTR_IMAGES,
    ATTR_PARAGRAPH_FORMAT,
    ATTR_REVISION,
    ATTR_SECTION,
    ATTR_STYLE,
    ATTR_TABLE,
    ATTR_TEXT,
    Change,
)
from ..core.package import diff_packages


def _paragraph_attrs(loc: str, old: dict, new: dict) -> list[Change]:
    """Every comparable property of one aligned pair of blocks."""
    changes: list[Change] = []
    if old.get("text") != new.get("text"):
        changes.append(Change(loc, ATTR_TEXT, old=old.get("text"), new=new.get("text")))
    if old.get("style") != new.get("style"):
        changes.append(Change(loc, ATTR_STYLE, old=old.get("style"),
                              new=new.get("style")))
    if old.get("format") != new.get("format"):
        changes.append(Change(loc, ATTR_FORMAT, detail="run formatting changed"))
    if old.get("properties") != new.get("properties"):
        changes.append(Change(loc, ATTR_PARAGRAPH_FORMAT,
                              detail="paragraph formatting changed"))
    if old.get("fields") != new.get("fields"):
        changes.append(Change(loc, ATTR_FIELD, old=old.get("fields"),
                              new=new.get("fields"), detail="field code changed"))
    if old.get("bookmarks") != new.get("bookmarks"):
        changes.append(Change(loc, ATTR_BOOKMARK, old=old.get("bookmarks"),
                              new=new.get("bookmarks"), detail="bookmarks changed"))
    if old.get("revisions") != new.get("revisions"):
        changes.append(Change(loc, ATTR_REVISION, old=old.get("revisions"),
                              new=new.get("revisions"),
                              detail="tracked-change markup changed"))
    if old.get("control") != new.get("control"):
        changes.append(Change(loc, ATTR_CONTENT_CONTROL, old=old.get("control"),
                              new=new.get("control"),
                              detail="content control properties changed"))
    return changes


def _table_attrs(loc: str, old: dict, new: dict) -> list[Change]:
    """Style, shape, geometry and cell contents of one pair of tables."""
    changes: list[Change] = []
    if old.get("style") != new.get("style"):
        changes.append(Change(loc, ATTR_TABLE, old=old.get("style"),
                              new=new.get("style"), detail="table style changed"))
    if old.get("geometry") != new.get("geometry"):
        changes.append(Change(loc, ATTR_TABLE,
                              detail="table widths, borders or shading changed"))
    o_rows, n_rows = old.get("rows", []), new.get("rows", [])
    o_shape = (len(o_rows), len(o_rows[0]) if o_rows else 0)
    n_shape = (len(n_rows), len(n_rows[0]) if n_rows else 0)
    if o_shape != n_shape:
        changes.append(Change(loc, ATTR_TABLE, old=o_shape, new=n_shape,
                              detail="table dimensions changed"))
    o_fmts, n_fmts = old.get("cell_formats", []), new.get("cell_formats", [])
    o_geom, n_geom = old.get("cell_geometry", []), new.get("cell_geometry", [])

    def at(grid, r, c):
        return grid[r][c] if r < len(grid) and c < len(grid[r]) else None

    for r in range(min(len(o_rows), len(n_rows))):
        for c in range(min(len(o_rows[r]), len(n_rows[r]))):
            cell_loc = f"{loc}!r{r + 1}c{c + 1}"
            if o_rows[r][c] != n_rows[r][c]:
                changes.append(Change(cell_loc, ATTR_TEXT,
                                      old=o_rows[r][c], new=n_rows[r][c]))
            if at(o_fmts, r, c) != at(n_fmts, r, c):
                changes.append(Change(cell_loc, ATTR_FORMAT,
                                      detail="run formatting changed"))
            if at(o_geom, r, c) != at(n_geom, r, c):
                changes.append(Change(cell_loc, ATTR_TABLE,
                                      detail="cell width, borders, shading or merge changed"))
    return changes


def _diff_blocks(base: dict, cand: dict) -> list[Change]:
    """Content controls, text boxes and nested tables, matched by location."""
    b_blocks, c_blocks = base.get("blocks", {}), cand.get("blocks", {})
    changes: list[Change] = []
    for location in sorted(b_blocks.keys() | c_blocks.keys()):
        old, new = b_blocks.get(location), c_blocks.get(location)
        if old is None or new is None:
            kind = (old or new).get("kind", "block")
            changes.append(Change(location, ATTR_TEXT,
                                  old=None if old is None else old.get("text"),
                                  new=None if new is None else new.get("text"),
                                  detail=f"{kind} {'added' if old is None else 'removed'}"))
            continue
        if old.get("kind") == "table" or new.get("kind") == "table":
            changes.extend(_table_attrs(location, old, new))
            continue
        changes.extend(_paragraph_attrs(location, old, new))
    return changes


def _diff_paragraphs_positional(b_paras: list, c_paras: list) -> list[Change]:
    changes: list[Change] = []
    for i in range(max(len(b_paras), len(c_paras))):
        loc = f"p{i + 1}"
        old = b_paras[i] if i < len(b_paras) else None
        new = c_paras[i] if i < len(c_paras) else None
        if old is None:
            changes.append(Change(loc, ATTR_TEXT, old=None, new=new["text"],
                                  detail="paragraph added"))
            continue
        if new is None:
            changes.append(Change(loc, ATTR_TEXT, old=old["text"], new=None,
                                  detail="paragraph removed"))
            continue
        changes.extend(_paragraph_attrs(loc, old, new))
    return changes


def _diff_paragraphs_aligned(b_paras: list, c_paras: list) -> list[Change]:
    """Align paragraphs by text, so an insertion does not shift everything.

    Added paragraphs are reported at their candidate position, removed ones at
    their baseline position; paragraphs that only moved report nothing.
    """
    matcher = SequenceMatcher(
        a=[p["text"] for p in b_paras], b=[p["text"] for p in c_paras], autojunk=False
    )
    changes: list[Change] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                changes.extend(
                    _paragraph_attrs(f"p{j1 + k + 1}", b_paras[i1 + k], c_paras[j1 + k])
                )
            continue
        # "replace", "delete" and "insert" all reduce to pairing what overlaps
        # and reporting the remainder as added or removed.
        for k in range(max(i2 - i1, j2 - j1)):
            old = b_paras[i1 + k] if i1 + k < i2 else None
            new = c_paras[j1 + k] if j1 + k < j2 else None
            if old is None:
                changes.append(Change(f"p{j1 + k + 1}", ATTR_TEXT, old=None,
                                      new=new["text"], detail="paragraph added"))
            elif new is None:
                changes.append(Change(f"p{i1 + k + 1}", ATTR_TEXT, old=old["text"],
                                      new=None, detail="paragraph removed"))
            else:
                changes.extend(_paragraph_attrs(f"p{j1 + k + 1}", old, new))
    return changes


def diff_snapshots(base: dict, cand: dict, *, align: bool = False) -> list[Change]:
    changes: list[Change] = []

    b_paras, c_paras = base["paragraphs"], cand["paragraphs"]
    if align:
        changes.extend(_diff_paragraphs_aligned(b_paras, c_paras))
    else:
        changes.extend(_diff_paragraphs_positional(b_paras, c_paras))

    b_tables, c_tables = base["tables"], cand["tables"]
    for i in range(max(len(b_tables), len(c_tables))):
        loc = f"table{i + 1}"
        old = b_tables[i] if i < len(b_tables) else None
        new = c_tables[i] if i < len(c_tables) else None
        if old is None or new is None:
            changes.append(Change(loc, ATTR_TABLE,
                                  old="present" if old else None,
                                  new="present" if new else None,
                                  detail="table added" if old is None else "table removed"))
            continue
        changes.extend(_table_attrs(loc, old, new))

    changes.extend(_diff_blocks(base, cand))

    b_secs, c_secs = base["sections"], cand["sections"]
    for i in range(max(len(b_secs), len(c_secs))):
        loc = f"section{i + 1}"
        old = b_secs[i] if i < len(b_secs) else None
        new = c_secs[i] if i < len(c_secs) else None
        if old != new:
            changes.append(Change(loc, ATTR_SECTION, old=old, new=new,
                                  detail="section/page setup changed"))

    b_hf, c_hf = base["header_footer"], cand["header_footer"]
    for i in range(max(len(b_hf), len(c_hf))):
        old = b_hf[i] if i < len(b_hf) else None
        new = c_hf[i] if i < len(c_hf) else None
        if old != new:
            changes.append(Change(f"section{i + 1}#header_footer", ATTR_HEADER_FOOTER,
                                  old=old, new=new))

    b_imgs, c_imgs = Counter(base["images"]), Counter(cand["images"])
    for sha in (b_imgs - c_imgs):
        changes.append(Change(f"#image:{sha[:8]}", ATTR_IMAGES, old="present", new=None,
                              detail="image removed or replaced"))
    for sha in (c_imgs - b_imgs):
        changes.append(Change(f"#image:{sha[:8]}", ATTR_IMAGES, old=None, new="present",
                              detail="image added or replaced"))

    # Comments, VBA, embedded objects and custom XML are dropped wholesale by
    # python-docx, so they are compared straight from the zip.
    changes.extend(diff_packages(base, cand))
    return changes
