"""OOXML package-part extraction, read straight from the zip container.

An .xlsx/.docx is a zip of XML and binary *parts*.  Editing libraries only
understand a subset of them: openpyxl silently discards charts, pivot tables,
shapes, embedded objects and the VBA project on save, and python-docx drops
comments and VBA.  Anything the library cannot model simply vanishes, and no
amount of inspecting the library's own object graph will reveal it — the part
is gone from both sides of that view.

So this layer never asks openpyxl or python-docx what is in the file.  It
opens the zip, hashes the parts it cares about, and compares the two
inventories.  A part that disappeared between baseline and candidate is
round-trip damage, whether or not any library can parse it.
"""

from __future__ import annotations

import hashlib
import zipfile
import xml.etree.ElementTree as ElementTree
from pathlib import Path

from .model import (
    ATTR_CHARTS,
    ATTR_COMMENTS,
    ATTR_CUSTOM_XML,
    ATTR_DRAWINGS,
    ATTR_EMBEDDED,
    ATTR_PIVOT_TABLES,
    ATTR_VBA,
    Change,
)

# category -> the part-name prefixes that belong to it.  Order matters: the
# first category whose prefix matches claims the part.
CATEGORY_PREFIXES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("vba", ("xl/vbaProject", "word/vbaProject")),
    ("charts", ("xl/charts/", "word/charts/")),
    ("pivot_tables", ("xl/pivotTables/", "xl/pivotCache/")),
    ("comments", ("xl/comments", "xl/threadedComments/", "word/comments")),
    ("embedded", ("xl/embeddings/", "word/embeddings/")),
    ("custom_xml", ("customXml/",)),
    ("drawings", ("xl/drawings/", "word/drawings/")),
)

# category -> the change attribute a policy addresses it by.
CATEGORY_ATTRIBUTES = {
    "vba": ATTR_VBA,
    "charts": ATTR_CHARTS,
    "pivot_tables": ATTR_PIVOT_TABLES,
    "comments": ATTR_COMMENTS,
    "embedded": ATTR_EMBEDDED,
    "custom_xml": ATTR_CUSTOM_XML,
    "drawings": ATTR_DRAWINGS,
}

CATEGORIES = tuple(name for name, _ in CATEGORY_PREFIXES)

# Parts that churn on any legitimate save and would otherwise cry wolf:
# document properties carry timestamps, the formula calc chain is rebuilt,
# printer settings are opaque binaries, and media is already compared
# semantically by the images attribute.
EXCLUDED_PREFIXES = (
    "docProps/",
    "xl/calcChain.xml",
    "xl/printerSettings/",
    "word/printerSettings/",
    "xl/media/",
    "word/media/",
    "[Content_Types].xml",
)

# Relationship files are rewritten wholesale (ids renumbered) whenever a part
# is touched, so they say nothing useful about content.
_RELS_SEGMENT = "_rels"

# Drawing elements worth counting.  A drawing part is summarized by which
# shapes it holds rather than by its bytes, for two reasons: openpyxl injects
# default children such as <a:ln><a:prstDash val="solid"/></a:ln> on every
# save, and pictures are already compared semantically by the images
# attribute.  Counting only non-picture shapes keeps both out of the way, so
# what remains is exactly the textboxes, chart frames and connectors that an
# editing library destroys without saying so.
_SHAPE_TAGS = frozenset({"sp", "graphicFrame", "cxnSp", "grpSp"})


def _is_excluded(name: str) -> bool:
    if name.endswith("/"):
        return True
    if _RELS_SEGMENT in name.split("/"):
        return True
    return name.startswith(EXCLUDED_PREFIXES)


def _categorize(name: str) -> str | None:
    for category, prefixes in CATEGORY_PREFIXES:
        if name.startswith(prefixes):
            return category
    return None


def _shape_summary(data: bytes) -> str | None:
    """Which non-picture shapes a drawing part holds, and their text.

    Returns "" when the part carries nothing but pictures — the images
    attribute already covers those, so tracking the part too would report
    every added or deleted image twice.  Returns None when the part is not
    parseable XML, so the caller can fall back to hashing the raw bytes
    rather than silently comparing nothing.
    """
    try:
        root = ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return None
    counts: dict[str, int] = {}
    texts: list[str] = []
    for element in root.iter():
        local = element.tag.rsplit("}", 1)[-1]
        if local in _SHAPE_TAGS:
            counts[local] = counts.get(local, 0) + 1
        elif local == "t" and element.text:
            texts.append(element.text)
    if not counts and not texts:
        return ""
    return repr((sorted(counts.items()), sorted(texts)))


def _digest(category: str, name: str, data: bytes) -> str | None:
    """The comparable fingerprint of a part, or None if it is not worth tracking."""
    if category == "drawings" and name.endswith(".xml"):
        summary = _shape_summary(data)
        if summary == "":
            return None
        if summary is not None:
            return hashlib.sha256(summary.encode("utf-8")).hexdigest()
    return hashlib.sha256(data).hexdigest()


def take_package_snapshot(path: str | Path) -> dict[str, dict[str, str]]:
    """Inventory the package parts of a document, grouped by category."""
    package: dict[str, dict[str, str]] = {name: {} for name in CATEGORIES}
    with zipfile.ZipFile(path) as zf:
        for info in zf.infolist():
            name = info.filename
            if _is_excluded(name):
                continue
            category = _categorize(name)
            if category is None:
                continue
            digest = _digest(category, name, zf.read(name))
            if digest is not None:
                package[category][name] = digest
    return package


def _short(digest: str | None) -> str | None:
    """A digest prefix: enough to tell parts apart, short enough to read."""
    return digest[:12] if digest else digest


def _location(category: str, name: str) -> str:
    # The VBA project keeps its long-standing standalone location so existing
    # policies that select "vba" go on working.
    if category == "vba":
        return "vba"
    return f"package#{category}:{name}"


def diff_packages(base: dict, cand: dict) -> list[Change]:
    """Compare two package inventories into Change records."""
    b_pkg = base.get("package", {})
    c_pkg = cand.get("package", {})
    changes: list[Change] = []
    for category in CATEGORIES:
        attribute = CATEGORY_ATTRIBUTES[category]
        b_parts = b_pkg.get(category, {})
        c_parts = c_pkg.get(category, {})
        for name in sorted(b_parts.keys() | c_parts.keys()):
            old = b_parts.get(name)
            new = c_parts.get(name)
            if old == new:
                continue
            if new is None:
                detail = f"{category} part removed: {name}"
            elif old is None:
                detail = f"{category} part added: {name}"
            else:
                detail = f"{category} part modified: {name}"
            changes.append(Change(_location(category, name), attribute,
                                  old=_short(old), new=_short(new), detail=detail))
    return changes
