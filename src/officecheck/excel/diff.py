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
    ATTR_VBA,
    Change,
)

_DEFAULT_CELL = {"value": None, "formula": None, "format": None}


def diff_snapshots(base: dict, cand: dict) -> list[Change]:
    changes: list[Change] = []

    base_sheets, cand_sheets = base["sheets"], cand["sheets"]
    for name in base_sheets.keys() - cand_sheets.keys():
        changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                              old="present", new=None, detail="sheet removed"))
    for name in cand_sheets.keys() - base_sheets.keys():
        changes.append(Change(f"sheet:{name}", ATTR_SHEET_STRUCTURE,
                              old=None, new="present", detail="sheet added"))

    for name in sorted(base_sheets.keys() & cand_sheets.keys()):
        changes.extend(_diff_sheet(name, base_sheets[name], cand_sheets[name]))

    for name in base.get("defined_names", {}).keys() | cand.get("defined_names", {}).keys():
        old = base.get("defined_names", {}).get(name)
        new = cand.get("defined_names", {}).get(name)
        if old != new:
            changes.append(Change(f"name:{name}", ATTR_DEFINED_NAMES, old=old, new=new))

    if base.get("vba_sha256") != cand.get("vba_sha256"):
        changes.append(Change("vba", ATTR_VBA,
                              old=base.get("vba_sha256"), new=cand.get("vba_sha256"),
                              detail="VBA project changed"))
    return changes


def _diff_sheet(name: str, b: dict, c: dict) -> list[Change]:
    changes: list[Change] = []

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
        for rng in b[key].keys() | c[key].keys():
            old, new = b[key].get(rng), c[key].get(rng)
            if old != new:
                changes.append(Change(f"{name}!{rng.split()[0]}", attr,
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
