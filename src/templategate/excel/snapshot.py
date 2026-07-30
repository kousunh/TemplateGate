"""Excel (.xlsx / .xlsm) snapshot extraction.

A snapshot is a plain dict of everything TemplateGate compares:
values, formulas, formats, merges, conditional formatting, data validation,
sheet structure, images, headers/footers, print settings, defined names and
the VBA project hash.  Extraction is read-only.
"""

from __future__ import annotations

import hashlib
from pathlib import Path

from openpyxl import load_workbook

from ..core.package import take_package_snapshot

_DEFAULT_CELL = {"value": None, "formula": None, "format": None}


def _color(c) -> tuple | None:
    if c is None:
        return None
    return (getattr(c, "rgb", None), getattr(c, "theme", None),
            getattr(c, "tint", None), getattr(c, "indexed", None))


def _side(s) -> tuple | None:
    if s is None or s.style is None:
        return None
    return (s.style, _color(s.color))


def _format_key(cell) -> tuple | None:
    """Normalize a cell's style into a comparable tuple. None == default."""
    if not cell.has_style:
        return None
    f, fill, b, a = cell.font, cell.fill, cell.border, cell.alignment
    return (
        ("font", f.name, f.size, bool(f.bold), bool(f.italic), f.underline,
         bool(f.strike), _color(f.color)),
        ("fill", fill.patternType, _color(getattr(fill, "fgColor", None)),
         _color(getattr(fill, "bgColor", None))),
        ("border", _side(b.left), _side(b.right), _side(b.top), _side(b.bottom),
         _side(b.diagonal)),
        ("align", a.horizontal, a.vertical, bool(a.wrap_text), a.text_rotation,
         a.indent),
        ("numfmt", cell.number_format),
        ("protect", cell.protection.locked, cell.protection.hidden),
    )


def _cells(ws_formula, ws_value) -> dict:
    cells: dict[str, dict] = {}
    for row in ws_formula.iter_rows():
        for cell in row:
            formula = None
            if isinstance(cell.value, str) and cell.value.startswith("="):
                formula = cell.value
            fmt = _format_key(cell)
            cached = ws_value[cell.coordinate].value if formula else cell.value
            if cached is None and formula is None and fmt is None:
                continue
            cells[cell.coordinate] = {
                "value": cached,
                "formula": formula,
                "format": fmt,
            }
    return cells


def _conditional_formatting(ws) -> dict:
    out: dict[str, list] = {}
    for cf in ws.conditional_formatting:
        rules = []
        for rule in cf.rules:
            rules.append((
                rule.type,
                getattr(rule, "operator", None),
                tuple(getattr(rule, "formula", None) or []),
                getattr(rule, "text", None),
                bool(getattr(rule, "stopIfTrue", False)),
            ))
        out[str(cf.sqref)] = sorted(map(repr, rules))
    return out


def _data_validation(ws) -> dict:
    out: dict[str, tuple] = {}
    for dv in ws.data_validations.dataValidation:
        out[str(dv.sqref)] = (
            dv.type, dv.operator, dv.formula1, dv.formula2,
            bool(dv.allowBlank), dv.showDropDown,
        )
    return out


def _images(ws) -> list[dict]:
    images = []
    for img in getattr(ws, "_images", []):
        try:
            data = img._data()
        except Exception:
            data = b""
        anchor = getattr(img.anchor, "_from", None)
        images.append({
            "sha256": hashlib.sha256(data).hexdigest(),
            "anchor": (anchor.col, anchor.row) if anchor is not None else None,
            "size": (round(img.width or 0), round(img.height or 0)),
        })
    return sorted(images, key=lambda d: (d["sha256"], d["anchor"] or (-1, -1)))


def _header_footer(ws) -> dict:
    out = {}
    for name in ("oddHeader", "oddFooter", "evenHeader", "evenFooter",
                 "firstHeader", "firstFooter"):
        hf = getattr(ws, name, None)
        if hf is None:
            continue
        parts = {side: getattr(hf, side).text for side in ("left", "center", "right")
                 if getattr(hf, side).text}
        if parts:
            out[name] = parts
    return out


def _print_settings(ws) -> dict:
    ps, pm = ws.page_setup, ws.page_margins
    return {
        "orientation": ps.orientation,
        "paper_size": ps.paperSize,
        "scale": ps.scale,
        "fit_to_width": ps.fitToWidth,
        "fit_to_height": ps.fitToHeight,
        "print_area": str(ws.print_area) if ws.print_area else None,
        "print_title_rows": ws.print_title_rows,
        "print_title_cols": ws.print_title_cols,
        "margins": (pm.left, pm.right, pm.top, pm.bottom, pm.header, pm.footer),
    }


def take_snapshot(path: str | Path) -> dict:
    path = Path(path)
    wb_formula = load_workbook(path, data_only=False)
    wb_value = load_workbook(path, data_only=True)

    sheets: dict[str, dict] = {}
    for index, name in enumerate(wb_formula.sheetnames):
        ws = wb_formula[name]
        sheets[name] = {
            "index": index,
            "visibility": ws.sheet_state,
            "cells": _cells(ws, wb_value[name]),
            "merges": sorted(str(m) for m in ws.merged_cells.ranges),
            "conditional_formatting": _conditional_formatting(ws),
            "data_validation": _data_validation(ws),
            "images": _images(ws),
            "header_footer": _header_footer(ws),
            "print": _print_settings(ws),
        }

    defined_names = {}
    for name, dn in wb_formula.defined_names.items():
        defined_names[name] = dn.value

    return {
        "target": "excel",
        "format": path.suffix.lstrip(".").lower(),
        "sheets": sheets,
        "defined_names": defined_names,
        "package": take_package_snapshot(path),
    }
