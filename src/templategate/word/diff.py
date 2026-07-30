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
    ATTR_FORMAT,
    ATTR_HEADER_FOOTER,
    ATTR_IMAGES,
    ATTR_SECTION,
    ATTR_STYLE,
    ATTR_TABLE,
    ATTR_TEXT,
    Change,
)
from ..core.package import diff_packages


def _paragraph_attrs(loc: str, old: dict, new: dict) -> list[Change]:
    """Text, style and direct-formatting differences of one aligned pair."""
    changes: list[Change] = []
    if old["text"] != new["text"]:
        changes.append(Change(loc, ATTR_TEXT, old=old["text"], new=new["text"]))
    if old["style"] != new["style"]:
        changes.append(Change(loc, ATTR_STYLE, old=old["style"], new=new["style"]))
    if old.get("format") != new.get("format"):
        changes.append(Change(loc, ATTR_FORMAT, detail="run formatting changed"))
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
        if old["style"] != new["style"]:
            changes.append(Change(loc, ATTR_TABLE, old=old["style"], new=new["style"],
                                  detail="table style changed"))
        o_rows, n_rows = old["rows"], new["rows"]
        o_shape = (len(o_rows), len(o_rows[0]) if o_rows else 0)
        n_shape = (len(n_rows), len(n_rows[0]) if n_rows else 0)
        if o_shape != n_shape:
            changes.append(Change(loc, ATTR_TABLE, old=o_shape, new=n_shape,
                                  detail="table dimensions changed"))
        o_fmts, n_fmts = old.get("cell_formats", []), new.get("cell_formats", [])
        for r in range(min(len(o_rows), len(n_rows))):
            for c in range(min(len(o_rows[r]), len(n_rows[r]))):
                cell_loc = f"{loc}!r{r + 1}c{c + 1}"
                if o_rows[r][c] != n_rows[r][c]:
                    changes.append(Change(cell_loc, ATTR_TEXT,
                                          old=o_rows[r][c], new=n_rows[r][c]))
                o_fmt = o_fmts[r][c] if r < len(o_fmts) and c < len(o_fmts[r]) else None
                n_fmt = n_fmts[r][c] if r < len(n_fmts) and c < len(n_fmts[r]) else None
                if o_fmt != n_fmt:
                    changes.append(Change(cell_loc, ATTR_FORMAT,
                                          detail="run formatting changed"))

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
