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
    ATTR_MARKUP,
    ATTR_MOVED,
    ATTR_PARAGRAPH_FORMAT,
    ATTR_REVISION,
    ATTR_SECTION,
    ATTR_STYLE,
    ATTR_TABLE,
    ATTR_TEXT,
    Change,
)
from ..core.describe import as_mapping, delta_values, describe_delta, field_delta
from ..core.package import diff_packages


def _merged_runs(formats) -> dict:
    """Fold a block's run formats into one property -> values map.

    A paragraph holds a set of run formats; naming the properties that differ
    across the whole set is what turns "run formatting changed" into
    "size '24' -> '8', color none -> 'FFFFFF'".
    """
    merged: dict[str, set] = {}
    for run in formats or ():
        for name, value in as_mapping(run).items():
            merged.setdefault(name, set()).add(repr(value))
    return {name: ", ".join(sorted(values)) for name, values in merged.items()}


def _breaks_text(breaks) -> str | None:
    if not breaks:
        return None
    return ", ".join(f"{count}x {kind}" for kind, count in breaks)


def _break_detail(old, new) -> str:
    before = dict(old or ())
    after = dict(new or ())
    parts = []
    for kind in sorted(before.keys() | after.keys()):
        was, now = before.get(kind, 0), after.get(kind, 0)
        if was == now:
            continue
        word = "page break" if kind == "page" else f"{kind} break"
        parts.append(f"{word} {'removed' if now < was else 'added'}"
                     f" ({was} -> {now})")
    return ", ".join(parts) or "breaks changed"


def _markup_text(mine, theirs) -> str | None:
    """The residual fragments only this side has, briefly."""
    unique = [fragment for fragment in (mine or ()) if fragment not in (theirs or ())]
    if not unique:
        return None
    shown = "; ".join(fragment[:80] for fragment in unique[:2])
    return shown + (f" (+{len(unique) - 2} more)" if len(unique) > 2 else "")


def _paragraph_attrs(loc: str, old: dict, new: dict) -> list[Change]:
    """Every comparable property of one aligned pair of blocks."""
    changes: list[Change] = []
    if old.get("text") != new.get("text"):
        changes.append(Change(loc, ATTR_TEXT, old=old.get("text"), new=new.get("text")))
    if old.get("style") != new.get("style"):
        changes.append(Change(loc, ATTR_STYLE, old=old.get("style"),
                              new=new.get("style")))
    # Compared as merged properties, not as the raw set of run formats: adding
    # or removing a run without changing any property is not a formatting
    # change, and reporting one would say "{} -> {}".
    delta = field_delta(_merged_runs(old.get("format")),
                        _merged_runs(new.get("format")))
    if delta:
        old_fields, new_fields = delta_values(delta)
        changes.append(Change(loc, ATTR_FORMAT, old=old_fields, new=new_fields,
                              detail=describe_delta(delta, "run formatting changed")))
    if old.get("properties") != new.get("properties"):
        delta = field_delta(old.get("properties"), new.get("properties"))
        old_fields, new_fields = delta_values(delta)
        changes.append(Change(loc, ATTR_PARAGRAPH_FORMAT,
                              old=old_fields, new=new_fields,
                              detail=describe_delta(delta, "paragraph formatting changed")))
    if old.get("breaks") != new.get("breaks"):
        changes.append(Change(loc, ATTR_PARAGRAPH_FORMAT,
                              old=_breaks_text(old.get("breaks")),
                              new=_breaks_text(new.get("breaks")),
                              detail=_break_detail(old.get("breaks"),
                                                   new.get("breaks"))))
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
    if old.get("markup") != new.get("markup"):
        changes.append(Change(loc, ATTR_MARKUP,
                              old=_markup_text(old.get("markup"), new.get("markup")),
                              new=_markup_text(new.get("markup"), old.get("markup")),
                              detail="markup changed that no other check covers"))
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
        delta = field_delta(old.get("geometry"), new.get("geometry"))
        old_fields, new_fields = delta_values(delta)
        changes.append(Change(loc, ATTR_TABLE, old=old_fields, new=new_fields,
                              detail=describe_delta(
                                  delta, "table widths, borders or shading changed")))
    o_rows, n_rows = old.get("rows", []), new.get("rows", [])
    o_shape = (len(o_rows), len(o_rows[0]) if o_rows else 0)
    n_shape = (len(n_rows), len(n_rows[0]) if n_rows else 0)
    if o_shape != n_shape:
        changes.append(Change(loc, ATTR_TABLE, old=o_shape, new=n_shape,
                              detail="table dimensions changed"))
    o_fmts, n_fmts = old.get("cell_formats", []), new.get("cell_formats", [])
    o_geom, n_geom = old.get("cell_geometry", []), new.get("cell_geometry", [])
    o_mark, n_mark = old.get("cell_markup", []), new.get("cell_markup", [])

    def at(grid, r, c):
        return grid[r][c] if r < len(grid) and c < len(grid[r]) else None

    for r in range(min(len(o_rows), len(n_rows))):
        for c in range(min(len(o_rows[r]), len(n_rows[r]))):
            cell_loc = f"{loc}!r{r + 1}c{c + 1}"
            if o_rows[r][c] != n_rows[r][c]:
                changes.append(Change(cell_loc, ATTR_TEXT,
                                      old=o_rows[r][c], new=n_rows[r][c]))
            cell_delta = field_delta(_merged_runs(at(o_fmts, r, c)),
                                     _merged_runs(at(n_fmts, r, c)))
            if cell_delta:
                old_fields, new_fields = delta_values(cell_delta)
                changes.append(Change(cell_loc, ATTR_FORMAT,
                                      old=old_fields, new=new_fields,
                                      detail=describe_delta(cell_delta,
                                                            "run formatting changed")))
            o_markup, n_markup = at(o_mark, r, c), at(n_mark, r, c)
            if o_markup != n_markup:
                changes.append(Change(cell_loc, ATTR_MARKUP,
                                      old=_markup_text(o_markup, n_markup),
                                      new=_markup_text(n_markup, o_markup),
                                      detail="markup changed that no other check covers"))
            geometry_delta = field_delta(at(o_geom, r, c), at(n_geom, r, c))
            if geometry_delta:
                old_fields, new_fields = delta_values(geometry_delta)
                changes.append(Change(cell_loc, ATTR_TABLE,
                                      old=old_fields, new=new_fields,
                                      detail=describe_delta(
                                          geometry_delta, "table cell changed")))
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


def _alignment(b_paras: list, c_paras: list) -> SequenceMatcher:
    return SequenceMatcher(a=[p["text"] for p in b_paras],
                           b=[p["text"] for p in c_paras], autojunk=False)


def _count(number: int) -> str:
    return f"{number} paragraph" if number == 1 else f"{number} paragraphs"


def _shift_report(b_paras: list, c_paras: list) -> tuple[set[int], str]:
    """Which positions only moved, and a plain description of the real edit.

    Comparing by position is what an acceptance gate wants — a template's
    paragraph 12 is supposed to stay paragraph 12.  But when one paragraph is
    deleted from sixty, every later position differs and the report becomes
    fifty-six lines that hide the single edit that caused them.
    """
    edits: list[str] = []
    aligned: dict[int, int] = {}
    for tag, i1, i2, j1, j2 in _alignment(b_paras, c_paras).get_opcodes():
        if tag == "equal":
            for offset in range(i2 - i1):
                aligned[i1 + offset] = j1 + offset
        elif tag == "delete":
            edits.append(f"{_count(i2 - i1)} removed at p{i1 + 1}")
        elif tag == "insert":
            edits.append(f"{_count(j2 - j1)} added at p{j1 + 1}")
        else:
            edits.append(f"{_count(i2 - i1)} rewritten at p{i1 + 1}")
    shifted = {index for index, target in aligned.items() if index != target}
    return shifted, "; ".join(edits)


def _diff_paragraphs_positional(b_paras: list, c_paras: list) -> list[Change]:
    shifted, summary = _shift_report(b_paras, c_paras)
    group = summary if shifted and summary else ""
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
        block = _paragraph_attrs(loc, old, new)
        if group and i in shifted:
            for change in block:
                change.group = group
                change.detail = (f"content shifted here by an earlier edit "
                                 f"({summary})")
        changes.extend(block)
    return changes


def _detect_moves(b_paras: list, c_paras: list,
                  matcher: SequenceMatcher) -> list[tuple[int, int]]:
    """Paragraphs whose text survived intact but landed somewhere else.

    Aligning by content is what makes an insertion cheap to report, but it
    also means a pure reordering aligns to nothing and would otherwise be
    reported as a pile of unrelated rewrites — or, under a policy that allows
    body text, as nothing at all.  Swapping two clauses of a contract is not
    a text edit; it is a move, and it gets said so.
    """
    b_texts = [p["text"] for p in b_paras]
    c_texts = [p["text"] for p in c_paras]
    anchored_old: set[int] = set()
    anchored_new: set[int] = set()
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            anchored_old.update(range(i1, i2))
            anchored_new.update(range(j1, j2))

    available: dict[str, list[int]] = {}
    for index, text in enumerate(c_texts):
        if index not in anchored_new and text:
            available.setdefault(text, []).append(index)

    moves: list[tuple[int, int]] = []
    for index, text in enumerate(b_texts):
        if index in anchored_old or not text:
            continue
        candidates = available.get(text)
        if not candidates:
            continue
        if index in candidates:
            # Same text, same position: it stayed put and only failed to
            # anchor because its neighbours moved around it.
            candidates.remove(index)
            continue
        moves.append((index, candidates.pop(0)))
    return moves


def _diff_paragraphs_aligned(b_paras: list, c_paras: list) -> list[Change]:
    """Align paragraphs by text, so an insertion does not shift everything.

    Added paragraphs are reported at their candidate position, removed ones at
    their baseline position, and a paragraph that only changed place is
    reported as moved.
    """
    matcher = _alignment(b_paras, c_paras)
    moves = _detect_moves(b_paras, c_paras, matcher)
    moved_old = {old for old, _ in moves}
    moved_new = {new for _, new in moves}

    changes: list[Change] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                changes.extend(
                    _paragraph_attrs(f"p{j1 + k + 1}", b_paras[i1 + k], c_paras[j1 + k])
                )
            continue
        # "replace", "delete" and "insert" all reduce to pairing what overlaps
        # and reporting the remainder as added or removed.  Paragraphs that
        # merely moved are taken out first so they are not mistaken for one.
        olds = [i for i in range(i1, i2) if i not in moved_old]
        news = [j for j in range(j1, j2) if j not in moved_new]
        for k in range(max(len(olds), len(news))):
            old = b_paras[olds[k]] if k < len(olds) else None
            new = c_paras[news[k]] if k < len(news) else None
            if old is None:
                changes.append(Change(f"p{news[k] + 1}", ATTR_TEXT, old=None,
                                      new=new["text"], detail="paragraph added"))
            elif new is None:
                changes.append(Change(f"p{olds[k] + 1}", ATTR_TEXT, old=old["text"],
                                      new=None, detail="paragraph removed"))
            else:
                changes.extend(_paragraph_attrs(f"p{news[k] + 1}", old, new))

    for old_index, new_index in sorted(moves, key=lambda move: move[1]):
        changes.append(Change(
            f"p{new_index + 1}", ATTR_MOVED,
            old=f"p{old_index + 1}", new=f"p{new_index + 1}",
            detail=f"paragraph moved from p{old_index + 1} to p{new_index + 1}"))
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

    # Header and footer parts, compared by the role they fill rather than the
    # filename they happen to have.  A section whose header text already
    # reported a change does not also report its part digest — that would say
    # the same thing twice, the second time opaquely.
    b_parts = base.get("header_footer_parts", {})
    c_parts = cand.get("header_footer_parts", {})
    said_already = {change.location.split("#", 1)[0] for change in changes
                    if change.attribute == ATTR_HEADER_FOOTER}
    for location in sorted(b_parts.keys() | c_parts.keys()):
        old, new = b_parts.get(location), c_parts.get(location)
        if old == new or location.split("#", 1)[0] in said_already:
            continue
        role = location.split("#", 1)[1].replace(":", " on the ") + " page"
        if new is None:
            detail = f"the {role} is gone from this section"
        elif old is None:
            detail = f"a {role} was added to this section"
        else:
            detail = f"the {role} changed"
        changes.append(Change(location, ATTR_HEADER_FOOTER,
                              old=None if old is None else "present",
                              new=None if new is None else "present",
                              detail=detail))

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
