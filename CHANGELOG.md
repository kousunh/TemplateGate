# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-07-31

Initial release.

### Added

- `templategate check` — compare a baseline and a candidate document against a
  trusted policy. Exit codes: `0` PASS, `1` FAIL, `2` execution error.
- `templategate diff`, `templategate snapshot`, and `templategate init`.
- Default-deny policy evaluation with `allow` / `protect` rules and
  `structural` settings.
- Excel (`.xlsx` / `.xlsm`) checks: values, formulas, formats, merged cells,
  conditional formatting, data validation, sheet structure, defined names,
  images, headers and footers, print settings, VBA.
- Word (`.docx`) checks: paragraphs, tables, styles, sections, headers and
  footers, images.
- OOXML package-part detection, read directly from the document's zip
  container rather than through an editing library, so parts that a library
  cannot represent are still compared. Covers `charts`, `pivot_tables`,
  `drawings` (shapes and textboxes), `comments`, `embedded`, and
  `custom_xml`, each available both as a `structural` key (default `strict`,
  set to `ignore` to opt out) and as an attribute in `allow` / `protect`.
  Parts are addressable as `package#*`, `package#<category>:*`, or by exact
  part name. Drawings are compared by shape summary rather than raw bytes, so
  the presentation defaults an editing library injects on a legitimate save
  do not raise false alarms. The VBA project keeps its existing `vba`
  selector, attribute, and location.
- `parts` and `links` structural keys. `parts` is a catch-all covering every
  package part, including ones the tool does not model, so an unrecognised
  part going missing counts as damage rather than passing unnoticed. `links`
  compares the external targets of relationships, catching a hyperlink
  repointed to a different URL while its display text is unchanged.
  Locations: `package#parts:<partname>`, `package#parts:<partname>#extLst`,
  `package#links:<url>`.
- Excel attributes `layout` (row and column visibility and sizes, addressable
  with whole-row and whole-column selectors such as `Sheet1!1:10`),
  `protection` (`Sheet1#protection`), and `sheet_settings`
  (`Sheet1#settings`, `workbook#settings`), the last covering a workbook
  switched to manual calculation. Data-validation enforcement fields,
  conditional-format `dxf` styling, sheet-scoped defined names
  (`name:Sheet1!<name>`) and rich-text run formatting are compared under the
  existing attributes. A workbook saved from `.xlsm` to `.xlsx`, losing its
  VBA project, is reported at `workbook#format`.
- Word attributes `paragraph_format`, `field`, `bookmark`, `revision`,
  `content_control`, and `moved`, read by walking `document.xml` directly
  rather than through the editing library's object graph. Content controls
  and text boxes are addressed as `sdt<N>` and `textbox<N>`, and a table
  nested in a cell as `<cell>!table<N>`. Text inserted as a tracked change
  counts as displayed text, so an agent editing through tracked changes is
  fully visible.
- Excel-style quoting for sheet names containing `'`, `!` or `#`:
  `'Q1!Q4'!A1` addresses cell A1 of the sheet named `Q1!Q4`, which bare
  `Q1!Q4` would otherwise read as cell Q4 of a sheet named `Q1`.

### Changed

- A candidate that still opens as a package but has lost parts it referenced
  is now treated as damaged rather than unreadable: the check FAILs (exit 1)
  and names the missing parts in plain language, instead of exiting 2 as a
  tool error. Exit 2 now means only that a file will not open at all, or that
  the configuration is wrong. Neither `review_only` mode nor
  `structural: ignore` can bless a document that cannot be opened.
- Reports name the sub-key that changed — `font.bold True -> False` rather
  than "the format changed". Where one edit shifts everything after it, the
  text and Markdown reports collapse the knock-on changes into one line while
  the JSON report keeps every change, related by a new `group` field.
- Optional semantic review (`off` / `review` / `gate`) through a
  user-supplied command; no vendor is pinned and `off` performs no network
  calls.
- Text, JSON, and Markdown reporters.
- Python API: `check`, `diff`, `snapshot`, `detect_target`.
- GitHub Action in `action/` and an agent skill in
  `skills/office-document-regression/`.

[Unreleased]: https://github.com/kousunh/TemplateGate/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/kousunh/TemplateGate/releases/tag/v0.1.0
