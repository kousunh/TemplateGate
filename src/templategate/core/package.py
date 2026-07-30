"""OOXML package-part extraction, read straight from the zip container.

An .xlsx/.docx is a zip of XML and binary *parts*.  Editing libraries only
understand a subset of them: openpyxl silently discards charts, pivot tables,
shapes, Excel tables, external links and the VBA project on save, and
python-docx drops comments and VBA.  Anything the library cannot model simply
vanishes, and no amount of inspecting the library's own object graph will
reveal it — the part is gone from both sides of that view.

So this layer never asks openpyxl or python-docx what is in the file.  It
opens the zip and takes an inventory of *every* part, default-deny: a part
nobody recognized still gets hashed and reported, because the parts that are
not modelled are exactly the ones that go missing.

Three things stop that from crying wolf on a legitimate edit:

* Parts whose content is already compared semantically upstream — worksheet
  XML, workbook.xml, shared strings, styles, document.xml — are tracked for
  presence only.  Editing an allowed cell rewrites the sheet part by
  definition, and the cell-level diff has already said so precisely.
* Word XML is normalized before hashing (revision ids, proofing marks and
  insignificant whitespace churn on every save without meaning anything).
* Drawing parts are reduced to the shapes they hold, because openpyxl injects
  presentation defaults on every save.

Volatile parts that carry no meaning at all — timestamps, the calc chain,
printer settings — are excluded outright.
"""

from __future__ import annotations

import hashlib
from collections import Counter
import re
import zipfile
import xml.etree.ElementTree as ElementTree
from pathlib import Path

from .model import (
    ATTR_CHARTS,
    ATTR_COMMENTS,
    ATTR_CUSTOM_XML,
    ATTR_DRAWINGS,
    ATTR_EMBEDDED,
    ATTR_LINKS,
    ATTR_PARTS,
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

# Everything the list above did not claim lands here.  This is what makes the
# inventory default-deny rather than an allowlist.
FALLBACK_CATEGORY = "parts"

# External relationship targets are compared as their own map: a hyperlink
# retargeted to another host changes no part content at all.
LINKS_CATEGORY = "links"

CATEGORIES = tuple(name for name, _ in CATEGORY_PREFIXES) + (
    FALLBACK_CATEGORY, LINKS_CATEGORY,
)

CATEGORY_ATTRIBUTES = {
    "vba": ATTR_VBA,
    "charts": ATTR_CHARTS,
    "pivot_tables": ATTR_PIVOT_TABLES,
    "comments": ATTR_COMMENTS,
    "embedded": ATTR_EMBEDDED,
    "custom_xml": ATTR_CUSTOM_XML,
    "drawings": ATTR_DRAWINGS,
    FALLBACK_CATEGORY: ATTR_PARTS,
    LINKS_CATEGORY: ATTR_LINKS,
}

# Parts that churn on any save and mean nothing on their own: document
# properties carry timestamps, the calc chain is rebuilt from the formulas,
# printer settings are opaque device blobs, content types are derived from
# which parts exist (which this module already tracks), and media is compared
# semantically by the images attribute.  docProps/custom.xml is deliberately
# NOT here: it holds the values DOCPROPERTY fields display.
EXCLUDED_PREFIXES = (
    "docProps/core.xml",
    "docProps/app.xml",
    "docProps/thumbnail",
    "[Content_Types].xml",
    "xl/calcChain.xml",
    "xl/printerSettings/",
    "word/printerSettings/",
    "xl/media/",
    "word/media/",
    # Worksheets are the one part family that legitimately comes and goes, and
    # both their presence and their contents are already reported by name
    # ("sheet:Notes", "Notes!B2").  Tracking the files as well would restate
    # every added or deleted sheet as an opaque internal filename.  A worksheet
    # part that vanishes without the sheet list noticing makes the workbook
    # unopenable, which the degraded read in api.py reports from the raw
    # name list instead.
    "xl/worksheets/",
)

# Parts whose *content* is already compared, cell by cell and paragraph by
# paragraph, by the format-specific snapshot.  Hashing them as well would
# report every allowed edit twice — once precisely, once as an opaque blob.
# Their presence is still tracked: losing one outright is real damage.
PRESENCE_ONLY_PREFIXES = (
    "xl/workbook.xml",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
    "word/document.xml",
)
_PRESENT = "present"

# Relationship files are rewritten wholesale (ids renumbered) whenever a part
# is touched, so their bytes say nothing — but the external targets inside
# them do, and those are extracted separately.
_RELS_SEGMENT = "_rels"

# Word rewrites these on every save without changing the document: revision
# save ids, proofing error marks, and whitespace between elements.
_WORD_NOISE = (
    re.compile(rb'\sw:rsid[A-Za-z]*="[^"]*"'),
    re.compile(rb"<w:proofErr[^/>]*/>"),
)
_BETWEEN_TAGS = re.compile(rb">\s+<")

# Drawing elements worth counting.  A drawing part is summarized by which
# shapes it holds rather than by its bytes, for two reasons: openpyxl injects
# default children such as <a:ln><a:prstDash val="solid"/></a:ln> on every
# save, and pictures are already compared semantically by the images
# attribute.  Counting only non-picture shapes keeps both out of the way, so
# what remains is exactly the textboxes, chart frames and connectors that an
# editing library destroys without saying so.
_SHAPE_TAGS = frozenset({"sp", "graphicFrame", "cxnSp", "grpSp"})

_RELATIONSHIP_TAG = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"

# Parts that are not hashed as a whole, but whose <extLst> extension blocks
# are.  An extension list is by construction the part of the XML the base
# schema does not describe — newer Excel features such as x14 data validation
# (the modern dropdown), sparklines and slicers all live there, and openpyxl
# drops the whole block on save while faithfully rewriting everything around
# it.  Hashing just the extensions catches that without re-reporting the cell
# edits the rest of the part legitimately carries.
_EXTENSION_SCANNED_PREFIXES = ("xl/worksheets/", "xl/workbook.xml")
_EXTENSION_SUFFIX = "#extLst"


def _is_rels(name: str) -> bool:
    return _RELS_SEGMENT in name.split("/")


def _is_excluded(name: str) -> bool:
    return name.endswith("/") or name.startswith(EXCLUDED_PREFIXES)


def _categorize(name: str) -> str:
    for category, prefixes in CATEGORY_PREFIXES:
        if name.startswith(prefixes):
            return category
    return FALLBACK_CATEGORY


def _normalize_word_xml(data: bytes) -> bytes:
    for pattern in _WORD_NOISE:
        data = pattern.sub(b"", data)
    return _BETWEEN_TAGS.sub(b"><", data).strip()


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
    if name.startswith(PRESENCE_ONLY_PREFIXES):
        return _PRESENT
    if category == "drawings" and name.endswith(".xml"):
        summary = _shape_summary(data)
        if summary == "":
            return None
        if summary is not None:
            return hashlib.sha256(summary.encode("utf-8")).hexdigest()
    if name.startswith("word/") and name.endswith(".xml"):
        data = _normalize_word_xml(data)
    return hashlib.sha256(data).hexdigest()


def _extension_digest(data: bytes) -> str | None:
    """Fingerprint the <extLst> blocks of a part, or None if it has none."""
    try:
        root = ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return None
    blocks: list[str] = []
    for element in root.iter():
        if element.tag.rsplit("}", 1)[-1] != "extLst":
            continue
        raw = ElementTree.tostring(element)
        try:
            blocks.append(ElementTree.canonicalize(raw.decode("utf-8")))
        except Exception:  # pragma: no cover - canonicalize is best effort
            blocks.append(raw.decode("utf-8", "replace"))
    if not blocks:
        return None
    return hashlib.sha256("".join(sorted(blocks)).encode("utf-8")).hexdigest()


def _external_targets(data: bytes) -> list[str]:
    """The external targets declared by one .rels part."""
    try:
        root = ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return []
    return [
        element.get("Target", "")
        for element in root.iter(_RELATIONSHIP_TAG)
        if element.get("TargetMode") == "External" and element.get("Target")
    ]


def take_package_snapshot(path: str | Path) -> dict[str, dict]:
    """Inventory the package parts of a document, grouped by category.

    Relationship ids are deliberately not part of the link map: they are
    renumbered whenever a part is rewritten, so keying on them would report a
    change every time nothing happened.  Targets are keyed by URL instead.
    """
    package: dict[str, dict] = {name: {} for name in CATEGORIES}
    links: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as zf:
        for info in zf.infolist():
            name = info.filename
            if name.endswith("/"):
                continue
            if _is_rels(name):
                for target in _external_targets(zf.read(name)):
                    links.setdefault(target, []).append(name)
                continue
            data = None
            if name.startswith(_EXTENSION_SCANNED_PREFIXES):
                data = zf.read(name)
                extensions = _extension_digest(data)
                if extensions is not None:
                    package[FALLBACK_CATEGORY][name + _EXTENSION_SUFFIX] = extensions
            if _is_excluded(name):
                continue
            if data is None:
                data = zf.read(name)
            category = _categorize(name)
            digest = _digest(category, name, data)
            if digest is not None:
                package[category][name] = digest
    package[LINKS_CATEGORY] = {url: sorted(set(owners))
                               for url, owners in sorted(links.items())}
    return package


def _escapes_root(name: str) -> str | None:
    """Why a part name does not stay inside the package, or None."""
    unified = name.replace("\\", "/")
    if unified.startswith("/"):
        return "absolute path"
    if len(unified) > 1 and unified[1] == ":":
        return "drive letter"
    if any(segment == ".." for segment in unified.split("/")):
        return "parent-directory segment"
    return None


def package_problem(path: str | Path) -> str | None:
    """Why this container cannot be compared at all, or None.

    A zip may legally hold two members with the same name, and readers do not
    agree about which one wins — Excel and Word may open a different document
    than the one the gate inspected.  A file like that cannot be checked, only
    refused; guessing would mean signing off on a document nobody can pin down.
    """
    try:
        with zipfile.ZipFile(path) as zf:
            names = [info.filename for info in zf.infolist()
                     if not info.filename.endswith("/")]
    except Exception:
        return None  # not a readable container: reported elsewhere

    counts = Counter(names)
    duplicates = sorted(name for name, count in counts.items() if count > 1)
    if duplicates:
        first = duplicates[0]
        extra = (f" (and {len(duplicates) - 1} other duplicated parts)"
                 if len(duplicates) > 1 else "")
        return (f"package contains {counts[first]} copies of {first}{extra}; "
                "different readers may open different documents")

    for name in names:
        reason = _escapes_root(name)
        if reason is not None:
            return (f"package contains a part outside the package root "
                    f"({reason}): {name}")
    return None


def list_part_names(path: str | Path) -> list[str]:
    """Every part in the container, unfiltered.

    The curated inventory deliberately leaves parts out; this does not.  It is
    what a degraded read falls back on, where the goal is to name whatever
    went missing rather than to keep the report quiet.
    """
    with zipfile.ZipFile(path) as zf:
        return sorted(info.filename for info in zf.infolist()
                      if not info.filename.endswith("/"))


def diff_part_names(base: dict, cand: dict) -> list[Change]:
    """Report parts that vanished or appeared, by raw name."""
    b_names = set(base.get("part_names", []))
    c_names = set(cand.get("part_names", []))
    changes: list[Change] = []
    for name in sorted(b_names - c_names):
        changes.append(Change(f"package#{FALLBACK_CATEGORY}:{name}", ATTR_PARTS,
                              old="present", new=None,
                              detail=f"{name} is missing from the candidate"))
    for name in sorted(c_names - b_names):
        changes.append(Change(f"package#{FALLBACK_CATEGORY}:{name}", ATTR_PARTS,
                              old=None, new="present",
                              detail=f"{name} was added to the candidate"))
    return changes


def _short(digest: str | None) -> str | None:
    """A digest prefix: enough to tell parts apart, short enough to read."""
    if digest is None or digest == _PRESENT:
        return digest
    return digest[:12]


def _location(category: str, name: str) -> str:
    # The VBA project keeps its long-standing standalone location so existing
    # policies that select "vba" go on working.
    if category == "vba":
        return "vba"
    return f"package#{category}:{name}"


def _describe(category: str, name: str, old, new) -> str:
    what = "external link" if category == LINKS_CATEGORY else f"{category} part"
    if new is None:
        return f"{what} removed: {name}"
    if old is None:
        return f"{what} added: {name}"
    return f"{what} modified: {name}"


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
            if category == LINKS_CATEGORY:
                changes.append(Change(
                    _location(category, name), attribute,
                    old="present" if old is not None else None,
                    new="present" if new is not None else None,
                    detail=_describe(category, name, old, new),
                ))
                continue
            changes.append(Change(_location(category, name), attribute,
                                  old=_short(old), new=_short(new),
                                  detail=_describe(category, name, old, new)))
    return changes
