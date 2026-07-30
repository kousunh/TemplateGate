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

OOXML package parts (both targets):
    "package#*"                          any package part
    "package#charts:*"                   any part in one category
    "package#charts:xl/charts/chart1.xml"  one specific part

Locations are produced by the extractors (see excel/diff.py, word/diff.py and
core/package.py).
"""

from __future__ import annotations

import re

from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils import range_boundaries


def _cell_to_tuple(coord: str) -> tuple[int, int]:
    col, row = coordinate_from_string(coord)
    return column_index_from_string(col), row


# Sheet-level location kinds that follow the "#" separator.  Anything else
# after a "#" is part of the sheet name — "#" is legal in Excel sheet names.
_LOCATION_KINDS = ("print", "header_footer")

# Package-part locations live in their own namespace, matched before any
# sheet-name parsing so a sheet called "package" cannot collide with them.
_PACKAGE_PREFIX = "package#"

# Excel grid limits, used to close the open ends of whole-column ("B:B") and
# whole-row ("4:4") ranges, for which openpyxl reports None boundaries.
_MAX_COL = 16384
_MAX_ROW = 1048576


def _is_location_kind(rest: str) -> bool:
    return rest in _LOCATION_KINDS or rest.startswith("image:")


def _bounds(ref: str) -> tuple[int, int, int, int] | None:
    """A1-style range boundaries with whole-column/row open ends closed."""
    try:
        min_c, min_r, max_c, max_r = range_boundaries(ref)
    except (ValueError, TypeError):
        return None
    return (
        1 if min_c is None else min_c,
        1 if min_r is None else min_r,
        _MAX_COL if max_c is None else max_c,
        _MAX_ROW if max_r is None else max_r,
    )


def _split_sheet(location: str) -> tuple[str, str, str]:
    """Split "Sheet1!B2" / "Sheet1#print" into (sheet, separator, rest).

    Sheet names may themselves contain "!" and "#", so the split happens at
    the *last* separator, "#" only counts when what follows it is a known
    location kind, and "!" only when what follows it is a real A1 range.
    ``separator`` is "" for a bare sheet name.
    """
    bang = location.rfind("!")
    hash_ = location.rfind("#")
    if hash_ > bang and _is_location_kind(location[hash_ + 1:]):
        return location[:hash_].strip("'"), "#", location[hash_ + 1:]
    if bang != -1 and _bounds(location[bang + 1:]) is not None:
        return location[:bang].strip("'"), "!", location[bang + 1:]
    return location.strip("'"), "", ""


def _range_contains(outer: str, inner: str) -> bool:
    """True if A1-style range/cell `inner` lies fully inside range `outer`."""
    o = _bounds(outer)
    i = _bounds(inner)
    if o is None or i is None:
        return False
    o_min_c, o_min_r, o_max_c, o_max_r = o
    i_min_c, i_min_r, i_max_c, i_max_r = i
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

    # OOXML package parts: "package#<category>:<partname>".
    if selector.startswith(_PACKAGE_PREFIX) or location.startswith(_PACKAGE_PREFIX):
        if not (selector.startswith(_PACKAGE_PREFIX)
                and location.startswith(_PACKAGE_PREFIX)):
            return False
        sel_rest = selector[len(_PACKAGE_PREFIX):]
        loc_rest = location[len(_PACKAGE_PREFIX):]
        if sel_rest == "*":
            return True
        if sel_rest.endswith(":*"):
            return loc_rest.startswith(sel_rest[:-1])  # whole category
        return sel_rest == loc_rest

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
    sel_sheet, sel_sep, sel_rest = _split_sheet(selector)
    loc_sheet, loc_sep, loc_rest = _split_sheet(location)
    if sel_sheet != loc_sheet:
        return False
    if not sel_sep:
        return True  # whole-sheet selector
    if not loc_sep:
        return False
    # A ranged ("!") selector never matches a sheet-level ("#") location such
    # as print settings or images, and vice versa.
    if sel_sep != loc_sep:
        return False
    if sel_sep == "#":
        return sel_rest == loc_rest
    return _range_contains(sel_rest, loc_rest)


def match_attributes(rule_attributes: list[str], attribute: str) -> bool:
    return "*" in rule_attributes or attribute in rule_attributes
