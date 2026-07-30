"""Read Word content from document.xml directly, not from python-docx.

python-docx models a useful subset of a document and quietly ignores the
rest.  ``Document.paragraphs`` walks only the body's direct ``w:p`` children,
so text inside a content control, a text box, a nested table or a tracked
insertion is not in it — and ``Document.tables`` stops at the top level.  A
gate that trusts that object graph is blind in exactly the places an edit is
easiest to hide.

So content is read from the XML.  Two rules keep it honest:

* **Displayed text is what a reader sees.**  Text inside ``w:ins`` counts,
  because a tracked insertion is on the page; text inside ``w:del`` does not,
  because a tracked deletion is struck from it.  Marking a clause deleted is
  therefore a text change, which is the point.
* **Each container gets its own location.**  Body paragraphs keep ``p{n}``,
  so existing policies still mean what they meant, and a content control or
  text box is addressed as ``sdt{k}`` / ``textbox{k}``.  Their text is left
  out of the enclosing paragraph, so an allowed body edit is reported once
  and a rule that permits body text does not silently permit the contents of
  a locked content control.
"""

from __future__ import annotations

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Containers that carry their own location, and deleted text.  Descending
# into these while reading a paragraph would either double-report them or
# count text that is not on the page.
_OPAQUE = frozenset({"del", "txbxContent", "sdt", "tbl", "instrText"})

_RUN_PROPERTIES = (
    "b", "i", "u", "strike", "dstrike", "caps", "smallCaps", "vanish",
    "webHidden", "outline", "shadow", "emboss", "vertAlign",
)
_PARAGRAPH_FLAGS = (
    "keepNext", "keepLines", "pageBreakBefore", "widowControl",
    "suppressLineNumbers", "outlineLvl", "bidi",
)


def _tag(element) -> str:
    tag = element.tag
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _val(element, attribute: str = "val"):
    if element is None:
        return None
    return element.get(W + attribute)


def _named(element, prefix: str, names: tuple[str, ...]) -> dict[str, str]:
    """Attributes that are actually set, under dotted names.

    Sparse and named so a difference reports as "indent.left 0 -> 3600"
    rather than as two positional tuples the reader has to line up by eye.
    """
    if element is None:
        return {}
    found = {}
    for name in names:
        value = element.get(W + name)
        if value is not None:
            found[f"{prefix}.{name}"] = value
    return found


def displayed_text(element) -> str:
    """The text a reader sees, in document order."""
    parts: list[str] = []
    seen_paragraph = False

    def walk(node) -> None:
        nonlocal seen_paragraph
        for child in node:
            name = _tag(child)
            if name in _OPAQUE:
                continue
            if name == "t":
                parts.append(child.text or "")
            elif name == "tab":
                parts.append("\t")
            elif name in ("br", "cr"):
                parts.append("\n")
            elif name == "noBreakHyphen":
                parts.append("-")
            elif name == "p":
                # A container holding several paragraphs keeps them apart,
                # without inventing a trailing newline on a single one.
                if seen_paragraph:
                    parts.append("\n")
                seen_paragraph = True
                walk(child)
            else:
                walk(child)

    walk(element)
    return "".join(parts)


def _run_format(run_properties) -> tuple:
    """One run's direct formatting, as the properties it actually sets.

    Sparse on purpose: a run with no properties yields nothing rather than a
    row of Nones, so comparing two runs names the property that differs
    ("size 24 -> 8") instead of listing every property either of them could
    have had.
    """
    if run_properties is None:
        return ()
    fields: dict[str, str] = {}
    for name in _RUN_PROPERTIES:
        node = run_properties.find(W + name)
        if node is not None:
            # An empty w:b means on; w:b w:val="0" means off.
            fields[name] = _val(node) or "1"
    for name, label in (("sz", "size"), ("szCs", "size.complex"),
                        ("color", "color"), ("u", "underline"),
                        ("highlight", "highlight"), ("spacing", "spacing")):
        value = _val(run_properties.find(W + name))
        if value is not None:
            fields[label] = value
    for element_name, label, attributes in (
            ("rFonts", "font", ("ascii", "hAnsi", "eastAsia", "cs")),
            ("shd", "shading", ("val", "color", "fill"))):
        node = run_properties.find(W + element_name)
        if node is None:
            continue
        for attribute in attributes:
            value = node.get(W + attribute)
            if value is not None:
                fields[f"{label}.{attribute}"] = value
    return tuple(sorted(fields.items()))


def _run_formats(element) -> list[tuple]:
    """The distinct run formats used, ignoring how the text is split up."""
    seen = set()

    def walk(node) -> None:
        for child in node:
            name = _tag(child)
            if name in _OPAQUE:
                continue
            if name == "r":
                seen.add(_run_format(child.find(W + "rPr")))
            else:
                walk(child)

    walk(element)
    return sorted(seen)


def paragraph_properties(paragraph) -> tuple:
    """Alignment, indents, spacing, tab stops and list membership.

    Moving a tab stop or demoting a clause to a deeper list level rewrites
    what the page says without touching a character of text.
    """
    properties = paragraph.find(W + "pPr")
    if properties is None:
        return ()
    fields: dict[str, object] = {}
    alignment = _val(properties.find(W + "jc"))
    if alignment is not None:
        fields["alignment"] = alignment
    fields.update(_named(properties.find(W + "ind"), "indent",
                         ("left", "right", "firstLine", "hanging", "start", "end")))
    fields.update(_named(properties.find(W + "spacing"), "spacing",
                         ("before", "after", "line", "lineRule")))
    fields.update(_named(properties.find(W + "framePr"), "frame",
                         ("w", "h", "x", "y")))
    tabs = properties.find(W + "tabs")
    if tabs is not None:
        stops = tuple(sorted((_val(tab, "val"), _val(tab, "pos"))
                             for tab in tabs.findall(W + "tab")))
        if stops:
            fields["tabs"] = stops
    numbering = properties.find(W + "numPr")
    if numbering is not None:
        for child, label in (("numId", "list.id"), ("ilvl", "list.level")):
            value = _val(numbering.find(W + child))
            if value is not None:
                fields[label] = value
    for name in _PARAGRAPH_FLAGS:
        node = properties.find(W + name)
        if node is not None:
            fields[name] = _val(node) or "1"
    return tuple(sorted(fields.items()))


def field_codes(element) -> list[str]:
    """Field instructions: what a field computes, not what it last showed.

    A DATE or DOCPROPERTY field displays a cached result that lives in an
    ordinary run, so swapping the instruction behind an unchanged-looking
    value is invisible at the text level.
    """
    codes: list[str] = []
    instruction: list[str] = []

    def walk(node) -> None:
        for child in node:
            name = _tag(child)
            if name in ("txbxContent", "sdt", "tbl"):
                continue
            if name == "fldSimple":
                codes.append((child.get(W + "instr") or "").strip())
            elif name == "instrText":
                instruction.append(child.text or "")
            else:
                walk(child)

    walk(element)
    if instruction:
        codes.append("".join(instruction).strip())
    return [code for code in codes if code]


def bookmarks(element) -> list[str]:
    names: list[str] = []

    def walk(node) -> None:
        for child in node:
            if _tag(child) == "bookmarkStart":
                name = child.get(W + "name")
                # Word writes _GoBack on every save; it means nothing.
                if name and name != "_GoBack":
                    names.append(name)
            elif _tag(child) not in ("txbxContent", "sdt", "tbl"):
                walk(child)

    walk(element)
    return sorted(names)


def breaks(element) -> tuple:
    """Page and column breaks, which are layout the text cannot show."""
    counts: dict[str, int] = {}

    def walk(node) -> None:
        for child in node:
            name = _tag(child)
            if name in ("txbxContent", "sdt", "tbl"):
                continue
            if name == "br":
                counts[_val(child, "type") or "line"] = (
                    counts.get(_val(child, "type") or "line", 0) + 1)
            elif name == "lastRenderedPageBreak":
                continue
            else:
                walk(child)

    walk(element)
    return tuple(sorted(counts.items()))


def revision_marks(element) -> tuple[int, int]:
    """How much of this block is a tracked insertion or deletion."""
    inserted = deleted = 0

    def walk(node) -> None:
        nonlocal inserted, deleted
        for child in node:
            name = _tag(child)
            if name == "ins":
                inserted += 1
            elif name == "del":
                deleted += 1
            elif name in ("txbxContent", "sdt", "tbl"):
                continue
            walk(child)

    walk(element)
    return (inserted, deleted)


def block_of(element, style_names: dict[str, str]) -> dict:
    """The comparable content of one paragraph-like block."""
    properties = element.find(W + "pPr")
    style_id = _val(properties.find(W + "pStyle")) if properties is not None else None
    return {
        "text": displayed_text(element),
        "style": style_names.get(style_id, style_id),
        "format": _run_formats(element),
        "properties": paragraph_properties(element),
        "fields": field_codes(element),
        "bookmarks": bookmarks(element),
        "revisions": revision_marks(element),
        "breaks": breaks(element),
    }


def sdt_properties(sdt) -> tuple:
    """A content control's identity and whether it is locked.

    Removing the lock leaves every visible thing about the clause unchanged
    and makes it editable by anyone.
    """
    properties = sdt.find(W + "sdtPr")
    if properties is None:
        return ()
    fields: dict[str, object] = {}
    for child, label in (("alias", "alias"), ("tag", "tag"), ("lock", "lock")):
        value = _val(properties.find(W + child))
        if value is not None:
            fields[label] = value
    if properties.find(W + "showingPlcHdr") is not None:
        fields["placeholder"] = "1"
    kinds = tuple(sorted(
        _tag(child) for child in properties
        if _tag(child) in ("text", "richText", "date", "dropDownList",
                           "comboBox", "checkbox", "picture", "docPartObj")
    ))
    if kinds:
        fields["kind"] = kinds
    return tuple(sorted(fields.items()))


def table_geometry(table) -> tuple:
    """Widths, borders, shading and layout of a table and its grid."""
    properties = table.find(W + "tblPr")
    grid = table.find(W + "tblGrid")
    fields: dict[str, object] = {}
    if properties is not None:
        fields.update(_named(properties.find(W + "tblW"), "width", ("w", "type")))
        fields.update(_named(properties.find(W + "tblInd"), "indent", ("w", "type")))
        fields.update(_named(properties.find(W + "shd"), "shading",
                             ("val", "color", "fill")))
        layout = _val(properties.find(W + "tblLayout"), "type")
        if layout is not None:
            fields["layout"] = layout
        alignment = _val(properties.find(W + "jc"))
        if alignment is not None:
            fields["alignment"] = alignment
        borders = properties.find(W + "tblBorders")
        for side in (borders if borders is not None else []):
            fields.update(_named(side, f"border.{_tag(side)}", ("val", "sz", "color")))
    if grid is not None:
        columns = tuple(_val(column, "w") for column in grid.findall(W + "gridCol"))
        if columns:
            fields["grid"] = columns
    return tuple(sorted(fields.items()))


def _collect_containers(node, blocks: dict, counters: dict, styles: dict) -> None:
    """Register every content control and text box found below a node.

    Numbering follows document order, which keeps a location stable as long
    as the containers around it are — the same bargain paragraph numbers make.
    """
    for child in node:
        name = _tag(child)
        if name == "sdt":
            counters["sdt"] += 1
            location = f"sdt{counters['sdt']}"
            content = child.find(W + "sdtContent")
            block = block_of(content, styles) if content is not None else {}
            blocks[location] = {**block, "kind": "content_control",
                                "control": sdt_properties(child)}
            if content is not None:
                _collect_containers(content, blocks, counters, styles)
                _collect_tables(content, location, blocks, counters, styles)
        elif name == "txbxContent":
            counters["textbox"] += 1
            location = f"textbox{counters['textbox']}"
            blocks[location] = {**block_of(child, styles), "kind": "textbox"}
            _collect_containers(child, blocks, counters, styles)
            _collect_tables(child, location, blocks, counters, styles)
        elif name != "tbl":
            _collect_containers(child, blocks, counters, styles)


def _collect_tables(node, prefix: str, blocks: dict, counters: dict,
                    styles: dict) -> None:
    """Register tables that are not top-level body tables."""
    for index, table in enumerate(
            [child for child in node if _tag(child) == "tbl"], start=1):
        location = f"{prefix}!table{index}"
        blocks[location] = {**table_of(table, location, blocks, counters, styles),
                            "kind": "table"}


def table_of(table, location: str, blocks: dict, counters: dict,
             styles: dict) -> dict:
    """One table, with its geometry and the containers inside its cells."""
    rows: list[list[str]] = []
    cell_formats: list[list] = []
    cell_geometries: list[list] = []
    table_properties = table.find(W + "tblPr")
    style_id = (_val(table_properties.find(W + "tblStyle"))
                if table_properties is not None else None)
    for row_index, row in enumerate(
            [child for child in table if _tag(child) == "tr"], start=1):
        cells = [child for child in row if _tag(child) == "tc"]
        rows.append([displayed_text(cell) for cell in cells])
        cell_formats.append([_run_formats(cell) for cell in cells])
        cell_geometries.append([cell_geometry(cell) for cell in cells])
        for cell_index, cell in enumerate(cells, start=1):
            cell_location = f"{location}!r{row_index}c{cell_index}"
            _collect_containers(cell, blocks, counters, styles)
            _collect_tables(cell, cell_location, blocks, counters, styles)
    return {
        "style": styles.get(style_id, style_id),
        "rows": rows,
        "cell_formats": cell_formats,
        "cell_geometry": cell_geometries,
        "geometry": table_geometry(table),
    }


def walk_body(body, styles: dict) -> tuple[list, list, dict]:
    """Split a document body into body paragraphs, tables and other containers.

    Body paragraphs keep the ``p{n}`` numbering they have always had — the
    walk counts the same direct ``w:p`` children python-docx did — so a policy
    written against p3 still means p3.
    """
    paragraphs: list[dict] = []
    tables: list[dict] = []
    blocks: dict[str, dict] = {}
    counters = {"sdt": 0, "textbox": 0}

    for child in body:
        name = _tag(child)
        if name == "p":
            paragraphs.append(block_of(child, styles))
            _collect_containers(child, blocks, counters, styles)
        elif name == "tbl":
            location = f"table{len(tables) + 1}"
            tables.append(table_of(table=child, location=location, blocks=blocks,
                                   counters=counters, styles=styles))
        elif name == "sdt":
            _collect_containers([child], blocks, counters, styles)
    return paragraphs, tables, blocks


def cell_geometry(cell) -> tuple:
    """Width, borders, shading, alignment and merge state of one cell."""
    properties = cell.find(W + "tcPr")
    if properties is None:
        return ()
    fields: dict[str, object] = {}
    fields.update(_named(properties.find(W + "tcW"), "width", ("w", "type")))
    fields.update(_named(properties.find(W + "shd"), "shading",
                         ("val", "color", "fill")))
    span = _val(properties.find(W + "gridSpan"))
    if span is not None:
        fields["span"] = span
    merge = properties.find(W + "vMerge")
    if merge is not None:
        fields["merged_vertically"] = _val(merge) or "continue"
    alignment = _val(properties.find(W + "vAlign"))
    if alignment is not None:
        fields["valign"] = alignment
    borders = properties.find(W + "tcBorders")
    for side in (borders if borders is not None else []):
        fields.update(_named(side, f"border.{_tag(side)}", ("val", "sz", "color")))
    return tuple(sorted(fields.items()))
