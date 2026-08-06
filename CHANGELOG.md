# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.3] - 2026-08-06

### Fixed

- `templategate --version` reported a stale number in both 0.1.1 (`0.1.0`)
  and 0.1.2 (`0.1.1`): a hardcoded `__version__` next to the real version in
  `pyproject.toml` let the two drift apart twice in a row. `__version__` is
  now read from the installed package's metadata at import time, so there is
  only one number to update per release.

## [0.1.2] - 2026-08-06

### Fixed

- `templategate --version` reported `0.1.0` in the 0.1.1 release: the
  package's `__version__` was not updated along with `pyproject.toml`.

## [0.1.1] - 2026-08-06

### Fixed

- A sheet that was renamed and, in the same edit, hidden, moved or swapped
  for a chartsheet reported only the rename. A renamed sheet's visibility,
  position and kind are now compared too, under its baseline name. Sheets
  that were merely paired for content comparison (a deletion plus an
  unrelated addition) still get no structure comparison.
- The CLI no longer surfaces openpyxl's "extension is not supported and
  will be removed" warning. The extension blocks it refers to (sparklines,
  x14 rules, slicers) are hashed and compared by the package layer, so the
  warning claimed the opposite of what happens. The library API leaves
  warning policy to the caller, as before.

### Changed

- The agent skill's trigger description now matches any edit to an existing
  Office document, in English and Japanese, instead of qualifying itself to
  documents "that must keep their structure intact".
- Both READMEs document skill installation per agent (Claude Code, Codex,
  ChatGPT / ChatGPT Work) and recommend committing the skill to the document
  repository's `.agents/skills/` next to the baseline and policy.

## [0.1.0] - 2026-07-31

Initial release.

### Added

- `templategate check` — compare a baseline and a candidate document against a
  trusted policy. Exit codes: `0` PASS, `1` FAIL, `2` execution error.
- `templategate diff`, `templategate snapshot`, and `templategate init`.
- `templategate suggest` — draft a commented policy from one reviewed edit.
  It sorts the differences into what the edit changed (allowed with tight
  selectors, contiguous cells merged into ranges), what the save produced
  (cached formula results, chart caches, a column width read back, a default
  font — allowed, each with the mechanism explained), and what neither
  accounts for (a replaced formula, a removed image, a dropped package part).
  The third group is never allowed automatically, so a draft taken from an
  edit that contained real damage names that damage in its header and does
  not pass the edit it came from. `--policy` proposes additions to an
  existing policy instead, commented out so that adopting one is a decision.
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
- Word attribute `markup`, a residual backstop comparing whatever is left in
  a block once every modelled attribute is accounted for. It needs no rule
  per feature, so it catches an unmodelled character style, a deleted
  footnote reference, a removed comment anchor, or a legacy form field
  disabled. Markup is held as a set of fragments, so re-flowing the same
  content across a different number of runs — what an allowed text edit
  routinely does — is not reported as a change.
- `sheet_settings` now covers `ignored_errors`, so suppressing Excel's
  warning triangles is visible.
- `--lang ja` on `check` and `diff` (or `TEMPLATEGATE_LANG=ja` set once for a
  whole job), for text and Markdown reports in Japanese — the register of
  README.ja.md, for the reader deciding whether a document may ship. The flag
  wins over the environment variable. The JSON report is unaffected by
  `--lang`: it stays English, since agents and CI parse it as a contract.

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
- A package whose member names are duplicated or which escape the package
  root is refused outright (exit 2) rather than compared, because different
  readers would open different documents from it. Damage remains a FAIL
  (exit 1); exit 2 now means the file could not be read unambiguously.

### Fixed

- A sheet with cells in far corners of the grid no longer hangs the check;
  cost tracks the number of populated cells rather than the addressed range.
- Optional semantic review (`off` / `review` / `gate`) through a
  user-supplied command; no vendor is pinned and `off` performs no network
  calls.
- Text, JSON, and Markdown reporters.
- Python API: `check`, `diff`, `snapshot`, `detect_target`.
- GitHub Action in `action/` and an agent skill in
  `skills/office-document-regression/`.

[Unreleased]: https://github.com/kousunh/TemplateGate/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/kousunh/TemplateGate/releases/tag/v0.1.0
