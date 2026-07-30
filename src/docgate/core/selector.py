"""Selector matching for policy rules.

Excel selectors:
    "*"                   any location
    "Sheet1"              anything on that sheet (cells, images, print, ...)
    "Sheet1!B2"           a single cell
    "Sheet1!B2:D100"      cells within a range
    "sheet:*"             sheet-structure locations
    "name:*"              defined-name locations
    "vba"                 the VBA project

Word selectors:
    "*"                   any location
    "body"                any paragraph (p<N>)
    "p3" / "p3-10"        paragraph index or range (1-based)
    "table2"              a whole table (structure and its cells)
    "table2!r1c2"         a single table cell
    "section1"            a section (page setup / header / footer)

Locations are produced by the extractors (see excel/diff.py and word/diff.py).
"""

from __future__ import annotations

import re

from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils import range_boundaries


def _cell_to_tuple(coord: str) -> tuple[int, int]:
    col, row = coordinate_from_string(coord)
    return column_index_from_string(col), row


def _split_sheet(location: str) -> tuple[str, str]:
    """Split "Sheet1!B2" / "Sheet1#print" into (sheet, rest). rest may be ""."""
    for sep in ("!", "#"):
        if sep in location:
            sheet, rest = location.split(sep, 1)
            return sheet.strip("'"), rest
    return location.strip("'"), ""


def _range_contains(outer: str, inner: str) -> bool:
    """True if A1-style range/cell `inner` lies fully inside range `outer`."""
    try:
        o_min_c, o_min_r, o_max_c, o_max_r = range_boundaries(outer)
        i_min_c, i_min_r, i_max_c, i_max_r = range_boundaries(inner)
    except ValueError:
        return False
    return (
        o_min_c <= i_min_c
        and o_min_r <= i_min_r
        and i_max_c <= o_max_c
        and i_max_r <= o_max_r
    )


def match_selector(selector: str, location: str) -> bool:
    selector = selector.strip()
    if selector == "*":
        return True

    # Special namespaces shared by both targets.
    for prefix in ("sheet:", "name:"):
        if selector.startswith(prefix):
            if not location.startswith(prefix):
                return False
            sel_rest = selector[len(prefix):]
            loc_rest = location[len(prefix):]
            return sel_rest == "*" or sel_rest == loc_rest
    if selector == "vba":
        return location == "vba"

    # Word-style selectors.
    if selector == "body":
        return bool(re.fullmatch(r"p\d+", location))
    m = re.fullmatch(r"p(\d+)(?:-(\d+))?", selector)
    if m:
        loc_m = re.fullmatch(r"p(\d+)", location)
        if not loc_m:
            return False
        idx = int(loc_m.group(1))
        lo = int(m.group(1))
        hi = int(m.group(2)) if m.group(2) else lo
        return lo <= idx <= hi
    m = re.fullmatch(r"(table\d+)(?:!(r\d+c\d+))?", selector)
    if m:
        table, cell = m.group(1), m.group(2)
        if cell:
            return location == f"{table}!{cell}"
        return location == table or location.startswith(f"{table}!")
    m = re.fullmatch(r"section\d+", selector)
    if m:
        return location == selector or location.startswith(f"{selector}#")

    # Excel-style selectors: "Sheet1" or "Sheet1!<range>".
    if location.startswith(("sheet:", "name:")) or location == "vba":
        return False
    sel_sheet, sel_rest = _split_sheet(selector)
    loc_sheet, loc_rest = _split_sheet(location)
    if sel_sheet != loc_sheet:
        return False
    if not sel_rest:
        return True  # whole-sheet selector
    if not loc_rest:
        return False
    # Both have a range/cell part; "#" locations (print, images...) never
    # match a ranged selector.
    if "#" in selector or "#" in location:
        return False
    return _range_contains(sel_rest, loc_rest)


def match_attributes(rule_attributes: list[str], attribute: str) -> bool:
    return "*" in rule_attributes or attribute in rule_attributes
