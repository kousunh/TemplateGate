"""Diff two Excel snapshots into Change records."""

from __future__ import annotations

from openpyxl.utils import get_column_letter

from ..core.model import (
    ATTR_CONDITIONAL_FORMATTING,
    ATTR_DATA_VALIDATION,
    ATTR_DEFINED_NAMES,
    ATTR_FORMAT,
    ATTR_FORMULA,
    ATTR_HEADER_FOOTER,
    ATTR_IMAGES,
    ATTR_LAYOUT,
    ATTR_MERGE,
    ATTR_PRINT_SETTINGS,
    ATTR_PROTECTION,
    ATTR_SHEET_SETTINGS,
    ATTR_SHEET_STRUCTURE,
    ATTR_VALUE,
    Change,
)
from ..core.describe import delta_values, describe_delta, field_delta
from ..core.messages import message
from ..core.package import diff_packages
from ..core.selector import quote_sheet

_DEFAULT_CELL = {"value": None, "formula": None, "format": None}


def _readable(value) -> str:
    """A cell value or formula as a person would write it, not as Python."""
    if value is None:
        return "empty"
    if isinstance(value, (list, tuple)):
        return " ".join(str(part) for part in value)
    return str(value)


def _similarity(b: dict, c: dict) -> float:
    """Fraction of cell coordinates that are byte-identical between two sheets."""
    b_cells, c_cells = b["cells"], c["cells"]
    keys = b_cells.keys() | c_cells.keys()
    if not keys:
        return 1.0
    same = sum(1 for k in keys if b_cells.get(k) == c_cells.get(k))
    return same / len(keys)


def _pair_renames(base_sheets: dict, cand_sheets: dict,
                  removed: list[str], added: list[str]) -> list[tuple[str, str, float]]:
    """Pair removed with added sheets, most-alike (then closest position) first.

    A rename must never hide the sheet's contents, so every removed sheet is
    paired while an added one is left — pairing is by best fit, not by a
    similarity threshold, because a rename that also rewrites every cell is
    exactly the case that has to stay visible.  The score rides along so the
    report can distinguish a real rename from two unrelated sheets that merely
    got paired up.
    """
    ranked = sorted(
        (
            (
                -_similarity(base_sheets[r], cand_sheets[a]),
                abs(base_sheets[r]["index"] - cand_sheets[a]["index"]),
                r,
                a,
            )
            for r in removed
            for a in added
        )
    )
    pairs: list[tuple[str, str, float]] = []
    used_r: set[str] = set()
    used_a: set[str] = set()
    for score, _distance, r, a in ranked:
        if r in used_r or a in used_a:
            continue
        used_r.add(r)
        used_a.add(a)
        pairs.append((r, a, -score))
    return pairs


# Above this share of identical cells a pairing is called a rename outright;
# below it the sheets are still compared, but the wording stays neutral.
_RENAME_SIMILARITY = 0.5


def _is_rename(pair: tuple[str, float] | None) -> bool:
    return pair is not None and pair[1] >= _RENAME_SIMILARITY


def _rename_detail(plain: str, renamed: str, pair: tuple[str, float] | None):
    if pair is None:
        return message(f"detail.sheet_{plain}")
    name, _score = pair
    if _is_rename(pair):
        return message(f"detail.sheet_renamed_{renamed}", name=name)
    # Two sheets that merely got paired up so their contents could still be
    # compared.  Calling that a rename would tell the reader a sheet survived
    # under a new name when it was in fact deleted.
    return message(f"detail.sheet_{plain}_paired", name=name)


def diff_snapshots(base: dict, cand: dict) -> list[Change]:
    changes: list[Change] = []

    base_sheets, cand_sheets = base["sheets"], cand["sheets"]
    removed = sorted(base_sheets.keys() - cand_sheets.keys())
    added = sorted(cand_sheets.keys() - base_sheets.keys())
    renames = _pair_renames(base_sheets, cand_sheets, removed, added)
    renamed_to = {r: (a, score) for r, a, score in renames}
    renamed_from = {a: (r, score) for r, a, score in renames}

    for name in removed:
        pair = renamed_to.get(name)
        changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                              old="present",
                              new=pair[0] if _is_rename(pair) else None,
                              detail=_rename_detail("removed", "to", pair)))
    for name in added:
        pair = renamed_from.get(name)
        changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                              old=pair[0] if _is_rename(pair) else None,
                              new="present",
                              detail=_rename_detail("added", "from", pair)))

    font_defaults = (dict(base.get("settings", ())), dict(cand.get("settings", ())))
    for name in sorted(base_sheets.keys() & cand_sheets.keys()):
        changes.extend(_diff_sheet(name, base_sheets[name], cand_sheets[name],
                                   font_defaults=font_defaults))

    # A renamed sheet keeps being compared cell by cell, reported under its
    # baseline name so that policy selectors written against the template
    # still apply.  Its position, kind and visibility are compared too, or a
    # sheet renamed *and* hidden in one edit reports only the rename.  A
    # low-similarity pairing gets no such comparison: the two sheets are
    # unrelated, and "sheet moved" between them would describe nothing.
    for old_name, new_name, score in renames:
        changes.extend(_diff_sheet(old_name, base_sheets[old_name],
                                   cand_sheets[new_name],
                                   include_structure=_is_rename((new_name, score)),
                                   font_defaults=font_defaults))

    for name in base.get("defined_names", {}).keys() | cand.get("defined_names", {}).keys():
        old = base.get("defined_names", {}).get(name)
        new = cand.get("defined_names", {}).get(name)
        if old != new:
            changes.append(Change(f"name:{name}", ATTR_DEFINED_NAMES, old=old, new=new))

    if base.get("format") != cand.get("format"):
        old_format, new_format = base.get("format"), cand.get("format")
        lost_vba = old_format == "xlsm" and new_format == "xlsx"
        changes.append(Change(
            "workbook#format", ATTR_SHEET_SETTINGS,
            old=old_format, new=new_format,
            detail=message("detail.format_changed_vba" if lost_vba
                           else "detail.format_changed",
                           old=old_format, new=new_format)))

    if base.get("settings", ()) != cand.get("settings", ()):
        delta = field_delta(base.get("settings"), cand.get("settings"))
        old_fields, new_fields = delta_values(delta)
        changes.append(Change("workbook#settings", ATTR_SHEET_SETTINGS,
                              old=old_fields, new=new_fields,
                              detail=describe_delta(delta, "detail.workbook_settings")))

    # Charts, pivot tables, shapes, VBA and friends live outside anything
    # openpyxl models, so they are compared straight from the zip.
    changes.extend(diff_packages(base, cand))
    return changes


def _at(anchor) -> str:
    """An image anchor as a cell reference a reader can find."""
    if not anchor:
        return "an unanchored position"
    column, row = anchor
    return f"{get_column_letter(int(column) + 1)}{int(row) + 1}"


# Office measures drawings in English Metric Units; 9525 of them make one
# screen pixel, which is the unit a person can picture.
_EMU_PER_PIXEL = 9525


def _size_text(size):
    """An image's on-page size, in pixels where the drawing gives us EMU.

    Returns a message rather than a string: this ends up inside the sentence
    about the image, and has to be spoken in the same language as that
    sentence rather than the one in force when the diff ran.
    """
    if not size:
        return message("size.unknown")
    kind = size[0]
    if kind == "emu":
        try:
            width, height = (round(int(value) / _EMU_PER_PIXEL) for value in size[1:3])
            return message("value.pixels", width=width, height=height)
        except (TypeError, ValueError):
            return message("size.unreadable")
    if kind == "to":
        return message("size.cell_span")
    return "x".join(str(part) for part in size)


def _diff_images(sheet: str, gone: set, arrived: set) -> list[Change]:
    """Image differences, saying *moved* when the picture itself is the same.

    The same picture at a new anchor arrives here as one removal and one
    addition carrying identical content hashes.  Reporting both is two lines
    that look like a replacement and read as if the picture might be gone;
    what actually happened is that someone nudged it.
    """
    changes: list[Change] = []
    before = {sha: (anchor, size) for sha, anchor, size in gone}
    after = {sha: (anchor, size) for sha, anchor, size in arrived}
    for sha in sorted(before.keys() & after.keys()):
        (old_anchor, old_size), (new_anchor, new_size) = before[sha], after[sha]
        moved, resized = old_anchor != new_anchor, old_size != new_size
        # One whole sentence per case rather than clauses joined with "and":
        # the joining word and the order of the clauses both move between
        # languages, so a stitched sentence is only ever right in one.
        if moved and resized:
            said = message("detail.image_moved_resized",
                          old_at=_at(old_anchor), new_at=_at(new_anchor),
                          old=_size_text(old_size), new=_size_text(new_size))
        elif moved:
            said = message("detail.image_moved",
                          old=_at(old_anchor), new=_at(new_anchor))
        elif resized:
            said = message("detail.image_resized",
                          old=_size_text(old_size), new=_size_text(new_size))
        else:
            said = message("detail.image_rewritten")
        changes.append(Change(
            f"{sheet}#image:{sha[:8]}", ATTR_IMAGES,
            old={"anchor": old_anchor, "size": old_size},
            new={"anchor": new_anchor, "size": new_size},
            detail=said))

    for sha in sorted(before.keys() - after.keys()):
        anchor, size = before[sha]
        changes.append(Change(f"{sheet}#image:{sha[:8]}", ATTR_IMAGES,
                              old={"anchor": anchor, "size": size}, new=None,
                              detail=message("detail.image_removed", where=_at(anchor))))
    for sha in sorted(after.keys() - before.keys()):
        anchor, size = after[sha]
        changes.append(Change(f"{sheet}#image:{sha[:8]}", ATTR_IMAGES,
                              old=None, new={"anchor": anchor, "size": size},
                              detail=message("detail.image_added", where=_at(anchor))))
    return changes


_INHERITED_FONT_FIELDS = (("font.name", "default_font.name"),
                          ("font.size", "default_font.size"))


def _resolve_inherited_fonts(delta: dict, defaults: tuple[dict, dict]) -> None:
    """Drop font differences that are only about who spelled the default out.

    A cell records a font only when it differs from its own workbook's
    default, and writers disagree about that default: ExcelJS resets it to
    Calibri while stamping the real font onto every cell, so the same-looking
    cell is "inherited" in one file and "explicit" in the other.  Comparing
    what each cell *renders as* settles it — and when both sides inherit, the
    difference is the workbook default itself, which is reported once at
    workbook#settings rather than on every cell.
    """
    base_defaults, cand_defaults = defaults
    for field, default_field in _INHERITED_FONT_FIELDS:
        if field not in delta:
            continue
        before, after = delta[field]
        if before is None and after is None:
            del delta[field]
            continue
        effective_before = before if before is not None else base_defaults.get(default_field)
        effective_after = after if after is not None else cand_defaults.get(default_field)
        if effective_before == effective_after:
            del delta[field]


def _diff_sheet(name: str, b: dict, c: dict, *,
                include_structure: bool = True,
                font_defaults: tuple[dict, dict] = ({}, {})) -> list[Change]:
    changes: list[Change] = []
    sheet = quote_sheet(name)

    if include_structure:
        if b.get("kind", "worksheet") != c.get("kind", "worksheet"):
            changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                                  old=b.get("kind"), new=c.get("kind"),
                                  detail=message("detail.sheet_kind")))
        if b["index"] != c["index"]:
            changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                                  old=b["index"], new=c["index"], detail=message("detail.sheet_moved")))
        if b["visibility"] != c["visibility"]:
            changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                                  old=b["visibility"], new=c["visibility"],
                                  detail=message("detail.sheet_visibility")))

    for coord in sorted(b["cells"].keys() | c["cells"].keys()):
        old = b["cells"].get(coord, _DEFAULT_CELL)
        new = c["cells"].get(coord, _DEFAULT_CELL)
        loc = f"{sheet}!{coord}"
        if old["value"] != new["value"]:
            # A cell whose formula did not move by one character, but whose
            # stored answer did, was recalculated rather than edited: Excel
            # recomputes on save, openpyxl throws the answer away instead.
            # Both are the saving tool talking, not the edit, so the change is
            # marked and the policy decides (``recalculation``).  A formula
            # that changed or vanished is never marked — that is the damage
            # this gate exists to catch, and it is reported under `formula`
            # as well as here.
            recalculated = bool(old["formula"]) and old["formula"] == new["formula"]
            changes.append(Change(loc, ATTR_VALUE, old=old["value"], new=new["value"],
                                  recalculated=recalculated))
        if old["formula"] != new["formula"]:
            said = ""
            if old["formula"] and not new["formula"]:
                # The archetypal damage: =SUM(B2:B4) becomes 1750.  It still
                # shows the right number today and stops being a total.
                said = message("detail.formula_hardcoded",
                               old=_readable(old["formula"]),
                               new=_readable(new["value"]))
            changes.append(Change(loc, ATTR_FORMULA, old=old["formula"],
                                  new=new["formula"], detail=said))
        if old["format"] != new["format"]:
            delta = field_delta(old["format"], new["format"])
            _resolve_inherited_fonts(delta, font_defaults)
            if not delta:
                continue
            if "numfmt" in delta:
                # An absent number format is not "no information", it is
                # General.  Printing it as "none" would read the same as a
                # format we merely failed to resolve, which is the one case
                # that matters most.
                before, after = delta["numfmt"]
                delta["numfmt"] = (before or "General", after or "General")
            old_fields, new_fields = delta_values(delta)
            changes.append(Change(loc, ATTR_FORMAT, old=old_fields, new=new_fields,
                                  detail=describe_delta(delta, "detail.cell_format")))

    for merged in sorted(set(b["merges"]) - set(c["merges"])):
        changes.append(Change(f"{sheet}!{merged}", ATTR_MERGE,
                              old="merged", new=None,
                              detail=message("detail.unmerged")))
    for merged in sorted(set(c["merges"]) - set(b["merges"])):
        changes.append(Change(f"{sheet}!{merged}", ATTR_MERGE,
                              old=None, new="merged",
                              detail=message("detail.merged")))

    for attr, key in ((ATTR_CONDITIONAL_FORMATTING, "conditional_formatting"),
                      (ATTR_DATA_VALIDATION, "data_validation")):
        for rng in sorted(b[key].keys() | c[key].keys()):
            old, new = b[key].get(rng), c[key].get(rng)
            if old != new:
                # A sqref may list several ranges ("B2:B5 D2:D5"); report each
                # one, or a rule covering an unmentioned range slips through.
                for part in str(rng).split() or [str(rng)]:
                    changes.append(Change(f"{sheet}!{part}", attr,
                                          old=_short(old), new=_short(new)))

    b_imgs = {(i["sha256"], tuple(i["anchor"]) if i["anchor"] else None,
               tuple(i["size"])) for i in b["images"]}
    c_imgs = {(i["sha256"], tuple(i["anchor"]) if i["anchor"] else None,
               tuple(i["size"])) for i in c["images"]}
    changes.extend(_diff_images(sheet, b_imgs - c_imgs, c_imgs - b_imgs))

    if b.get("protection", ()) != c.get("protection", ()):
        delta = field_delta(b.get("protection"), c.get("protection"))
        old_fields, new_fields = delta_values(delta)
        changes.append(Change(f"{sheet}#protection", ATTR_PROTECTION,
                              old=old_fields or _short(b.get("protection")),
                              new=new_fields or _short(c.get("protection")),
                              detail=describe_delta(delta, "detail.sheet_protection")))

    if b.get("settings", ()) != c.get("settings", ()):
        delta = field_delta(b.get("settings"), c.get("settings"))
        old_fields, new_fields = delta_values(delta)
        changes.append(Change(f"{sheet}#settings", ATTR_SHEET_SETTINGS,
                              old=old_fields, new=new_fields,
                              detail=describe_delta(delta, "detail.sheet_settings")))

    b_layout, c_layout = b.get("layout", {}), c.get("layout", {})
    for ref in sorted(b_layout.keys() | c_layout.keys()):
        old, new = b_layout.get(ref), c_layout.get(ref)
        if old == new:
            continue
        # A row or column that goes back to its default stops being recorded
        # at all.  Read as "absent", unhiding a column would report its width
        # vanishing rather than the thing that actually happened.
        old_map, new_map = dict(old or ()), dict(new or ())
        for side in (old_map, new_map):
            side.setdefault("hidden", False)
        delta = field_delta(old_map, new_map)
        old_fields, new_fields = delta_values(delta)
        said = describe_delta(delta, "detail.row_column")
        still_hidden = (old_map.get("hidden") and new_map.get("hidden")
                        and "hidden" not in delta)
        if still_hidden:
            said = message("detail.still_hidden", body=said)
        changes.append(Change(f"{sheet}!{ref}", ATTR_LAYOUT,
                              old=old_fields, new=new_fields, detail=said))

    b_names, c_names = b.get("defined_names", {}), c.get("defined_names", {})
    for dn in sorted(b_names.keys() | c_names.keys()):
        old, new = b_names.get(dn), c_names.get(dn)
        if old != new:
            changes.append(Change(f"name:{sheet}!{dn}", ATTR_DEFINED_NAMES,
                                  old=old, new=new,
                                  detail=message("detail.defined_name_sheet")))

    if b["header_footer"] != c["header_footer"]:
        changes.append(Change(f"{sheet}#header_footer", ATTR_HEADER_FOOTER,
                              old=b["header_footer"], new=c["header_footer"]))

    if b["print"] != c["print"]:
        fields = ", ".join(sorted(
            k for k in b["print"].keys() | c["print"].keys()
            if b["print"].get(k) != c["print"].get(k)
        ))
        changes.append(Change(f"{sheet}#print", ATTR_PRINT_SETTINGS,
                              detail=message("detail.print_settings_fields", fields=fields)))
    return changes


def _short(value, limit: int = 200):
    text = repr(value)
    return text if len(text) <= limit else text[:limit] + "..."
