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
- Optional semantic review (`off` / `review` / `gate`) through a
  user-supplied command; no vendor is pinned and `off` performs no network
  calls.
- Text, JSON, and Markdown reporters.
- Python API: `check`, `diff`, `snapshot`, `detect_target`.
- GitHub Action in `action/` and an agent skill in
  `skills/office-document-regression/`.

[Unreleased]: https://github.com/kousunh/TemplateGate/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/kousunh/TemplateGate/releases/tag/v0.1.0
