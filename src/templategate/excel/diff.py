"""Diff two Excel snapshots into Change records."""

from __future__ import annotations

from ..core.model import (
    ATTR_CONDITIONAL_FORMATTING,
    ATTR_DATA_VALIDATION,
    ATTR_DEFINED_NAMES,
    ATTR_FORMAT,
    ATTR_FORMULA,
    ATTR_HEADER_FOOTER,
    ATTR_IMAGES,
    ATTR_MERGE,
    ATTR_PRINT_SETTINGS,
    ATTR_SHEET_STRUCTURE,
    ATTR_VALUE,
    Change,
)
from ..core.package import diff_packages

_DEFAULT_CELL = {"value": None, "formula": None, "format": None}


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


def _rename_detail(plain: str, renamed: str, pair: tuple[str, float] | None) -> str:
    if pair is None:
        return f"sheet {plain}"
    name, score = pair
    if score >= _RENAME_SIMILARITY:
        return f"sheet {renamed} {name!r}"
    return f"sheet {plain}; contents compared against {name!r}"


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
                              old="present", new=pair[0] if pair else None,
                              detail=_rename_detail("removed", "renamed to", pair)))
    for name in added:
        pair = renamed_from.get(name)
        changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                              old=pair[0] if pair else None, new="present",
                              detail=_rename_detail("added", "renamed from", pair)))

    for name in sorted(base_sheets.keys() & cand_sheets.keys()):
        changes.extend(_diff_sheet(name, base_sheets[name], cand_sheets[name]))

    # A renamed sheet keeps being compared cell by cell, reported under its
    # baseline name so that policy selectors written against the template
    # still apply.
    for old_name, new_name, _score in renames:
        changes.extend(_diff_sheet(old_name, base_sheets[old_name],
                                   cand_sheets[new_name], include_structure=False))

    for name in base.get("defined_names", {}).keys() | cand.get("defined_names", {}).keys():
        old = base.get("defined_names", {}).get(name)
        new = cand.get("defined_names", {}).get(name)
        if old != new:
            changes.append(Change(f"name:{name}", ATTR_DEFINED_NAMES, old=old, new=new))

    # Charts, pivot tables, shapes, VBA and friends live outside anything
    # openpyxl models, so they are compared straight from the zip.
    changes.extend(diff_packages(base, cand))
    return changes


def _diff_sheet(name: str, b: dict, c: dict, *,
                include_structure: bool = True) -> list[Change]:
    changes: list[Change] = []

    if include_structure:
        if b["index"] != c["index"]:
            changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                                  old=b["index"], new=c["index"], detail="sheet moved"))
        if b["visibility"] != c["visibility"]:
            changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                                  old=b["visibility"], new=c["visibility"],
                                  detail="sheet visibility changed"))

    for coord in sorted(b["cells"].keys() | c["cells"].keys()):
        old = b["cells"].get(coord, _DEFAULT_CELL)
        new = c["cells"].get(coord, _DEFAULT_CELL)
        loc = f"{name}!{coord}"
        if old["value"] != new["value"]:
            changes.append(Change(loc, ATTR_VALUE, old=old["value"], new=new["value"]))
        if old["formula"] != new["formula"]:
            changes.append(Change(loc, ATTR_FORMULA, old=old["formula"], new=new["formula"]))
        if old["format"] != new["format"]:
            changes.append(Change(loc, ATTR_FORMAT, detail="cell format changed"))

    for merged in sorted(set(b["merges"]) - set(c["merges"])):
        changes.append(Change(f"{name}!{merged}", ATTR_MERGE,
                              old="merged", new=None, detail="merge removed"))
    for merged in sorted(set(c["merges"]) - set(b["merges"])):
        changes.append(Change(f"{name}!{merged}", ATTR_MERGE,
                              old=None, new="merged", detail="merge added"))

    for attr, key in ((ATTR_CONDITIONAL_FORMATTING, "conditional_formatting"),
                      (ATTR_DATA_VALIDATION, "data_validation")):
        for rng in sorted(b[key].keys() | c[key].keys()):
            old, new = b[key].get(rng), c[key].get(rng)
            if old != new:
                # A sqref may list several ranges ("B2:B5 D2:D5"); report each
                # one, or a rule covering an unmentioned range slips through.
                for part in str(rng).split() or [str(rng)]:
                    changes.append(Change(f"{name}!{part}", attr,
                                          old=_short(old), new=_short(new)))

    b_imgs = {(i["sha256"], tuple(i["anchor"]) if i["anchor"] else None,
               tuple(i["size"])) for i in b["images"]}
    c_imgs = {(i["sha256"], tuple(i["anchor"]) if i["anchor"] else None,
               tuple(i["size"])) for i in c["images"]}
    for sha, anchor, size in sorted(b_imgs - c_imgs):
        changes.append(Change(f"{name}#image:{sha[:8]}", ATTR_IMAGES,
                              old={"anchor": anchor, "size": size}, new=None,
                              detail="image removed, replaced, moved or resized"))
    for sha, anchor, size in sorted(c_imgs - b_imgs):
        changes.append(Change(f"{name}#image:{sha[:8]}", ATTR_IMAGES,
                              old=None, new={"anchor": anchor, "size": size},
                              detail="image added, replaced, moved or resized"))

    if b["header_footer"] != c["header_footer"]:
        changes.append(Change(f"{name}#header_footer", ATTR_HEADER_FOOTER,
                              old=b["header_footer"], new=c["header_footer"]))

    if b["print"] != c["print"]:
        detail = ", ".join(sorted(
            k for k in b["print"].keys() | c["print"].keys()
            if b["print"].get(k) != c["print"].get(k)
        ))
        changes.append(Change(f"{name}#print", ATTR_PRINT_SETTINGS,
                              detail=f"print settings changed: {detail}"))
    return changes


def _short(value, limit: int = 200):
    text = repr(value)
    return text if len(text) <= limit else text[:limit] + "..."
