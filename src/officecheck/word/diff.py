"""Diff two Word snapshots into Change records.

Paragraphs and tables are compared by position (1-based index).  Inserting a
paragraph in the middle therefore reports every following paragraph as
changed — deliberate for an acceptance gate: position shifts in a fixed
template are exactly what it must catch.  Policies for append-style edits can
allow a trailing paragraph range instead.
"""

from __future__ import annotations

from collections import Counter

from ..core.model import (
    ATTR_HEADER_FOOTER,
    ATTR_IMAGES,
    ATTR_SECTION,
    ATTR_STYLE,
    ATTR_TABLE,
    ATTR_TEXT,
    Change,
)


def diff_snapshots(base: dict, cand: dict) -> list[Change]:
    changes: list[Change] = []

    b_paras, c_paras = base["paragraphs"], cand["paragraphs"]
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
        if old["text"] != new["text"]:
            changes.append(Change(loc, ATTR_TEXT, old=old["text"], new=new["text"]))
        if old["style"] != new["style"]:
            changes.append(Change(loc, ATTR_STYLE, old=old["style"], new=new["style"]))

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
        for r in range(min(len(o_rows), len(n_rows))):
            for c in range(min(len(o_rows[r]), len(n_rows[r]))):
                if o_rows[r][c] != n_rows[r][c]:
                    changes.append(Change(f"{loc}!r{r + 1}c{c + 1}", ATTR_TEXT,
                                          old=o_rows[r][c], new=n_rows[r][c]))

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
    return changes
