# TemplateGate

**English** | [日本語](README.ja.md)

[![CI](https://github.com/kousunh/TemplateGate/actions/workflows/ci.yml/badge.svg)](https://github.com/kousunh/TemplateGate/actions/workflows/ci.yml)

**Policy-as-code acceptance gate for AI-edited Excel & Word documents.**

AI agents (Claude Code, Codex, ChatGPT, ...) are great at editing Excel and Word
files — and occasionally destroy a formula, a merged cell, or a print layout
while doing it. TemplateGate is a **read-only regression test** for
`.xlsx` / `.xlsm` / `.docx` files: it compares a *baseline* (before the edit)
with a *candidate* (after the edit) and verifies that **only the changes your
policy allows were made**.

- **Default deny / fail closed** — any change not explicitly allowed is a violation.
- **Deterministic structural checks** — values, formulas, formats, merged cells,
  conditional formatting, data validation, sheet structure, images (count /
  content hash / position / size), headers & footers, print settings, VBA,
  Word paragraphs / tables / sections.
- **Catches what editing tools silently throw away** — charts, pivot tables,
  comments, embedded objects, custom XML, Excel shapes & textboxes, and every
  other part of the file besides. Parts are read straight out of the document's
  internal package, so damage is caught even when the tool that made the edit
  could not represent those parts in the first place — the failure mode where
  you get a file back that opens fine and looks fine, minus its charts. A part
  nobody recognises still counts: unknown is not the same as allowed.
- **Catches what you cannot see** — a row quietly hidden, a sheet left
  unlocked, a workbook switched to manual calculation so every formula shows a
  stale answer, a hyperlink repointed while its display text stays put, body
  text set to 4pt white. Edits made as tracked changes are visible too: an
  insertion is text on the page, so it is reported as one.
- **Optional semantic checks** — `off` (default; nothing ever leaves your
  machine), `review` (AI findings as warnings), or `gate` (AI findings affect
  PASS/FAIL). Bring your own model — TemplateGate does not pin a vendor.
- **Not an editor** — TemplateGate never modifies documents and never
  auto-repairs. A failed candidate should be discarded; a human makes the
  final call.

## The trust boundary

The agent that edited the document must never be the one that decides what is
allowed. An agent may *propose* a policy, but the actual check runs against a
**trusted policy** that a human (or CI) has reviewed and pinned. TemplateGate
only reads the policy file — it has no way to widen it.

This is not a hypothetical risk. Microsoft Research measured frontier LLMs
corrupting an average of 25% of document content over long editing workflows —
"sparse but severe errors that silently corrupt documents" — and found that
agentic tool use alone does not fix it
([Laban, Schnabel & Neville, 2026](https://arxiv.org/abs/2604.15597)).
Verification has to live outside the agent.

## Install

```bash
pip install git+https://github.com/kousunh/TemplateGate.git
```

Publishing to PyPI is planned, after which `pip install templategate` will be
the install command. Until then, install from the repository.

Requires Python 3.10+. No Office installation and no network access are needed
to run a check. Dependencies: openpyxl, python-docx (which brings lxml),
Pillow, and PyYAML.

## Quick start

```bash
# 1. Generate a starter policy and edit it
templategate init --target excel

# 2. Let your agent edit a COPY of the document

# 3. Check the result
templategate check \
  --baseline plan_2026.xlsx \
  --candidate plan_2026.edited.xlsx \
  --policy templategate.policy.yaml \
  --report json
```

Exit codes: `0` = PASS, `1` = FAIL, `2` = execution error.

A candidate that still opens as a package but has lost parts it referenced is
**damaged, not unreadable** — that is a FAIL (exit 1) with the missing parts
named, not a tool error.

Exit 2 means the file could not be read unambiguously: it will not open at
all, or its package is ambiguous. A zip holding two copies of the same part is
refused rather than guessed at, because different readers would open different
documents from it. Configuration mistakes are exit 2 as well. Neither
`review_only` mode nor `structural: ignore` can bless a document that cannot
be read.

Between those two there is a middle state worth knowing about. When a file
comes from a writer that omits parts the usual readers expect — document
properties, for instance — the check falls back to comparing package parts
alone. It says so in the report rather than pretending otherwise. That is an
honest comparison but a much weaker one: part-level hashes will tell you a
sheet changed, not which cell. If you see it, prefer a writer that produces a
complete package.

A policy looks like this:

```yaml
version: 1
target: excel
allow:
  - selector: "Sheet1!B2:B100"   # the agent may change these values...
    attributes: [value]
protect:
  - selector: "*"                 # ...but formulas and print settings, never
    attributes: [formula, print_settings, vba]
structural:
  sheets: strict                  # adding/removing/hiding sheets fails
  images: strict
  charts: strict                  # so does losing a chart, a pivot table,
  pivot_tables: strict            # a comment, an embedded object, custom XML,
  comments: strict                # or an Excel shape / textbox
  parts: strict                   # ...or any other part of the file
  links: strict                   # or a hyperlink repointed elsewhere
semantic:
  mode: "off"                     # off | review | gate
```

Every structural category is `strict` unless you say otherwise — deleting the
line does not turn it off. Set one to `ignore` to opt out.

Two worked examples live in [`examples/`](examples/): an Excel policy that
lets an agent update quantity cells and nothing else
([excel-quantity-update.policy.yaml](examples/excel-quantity-update.policy.yaml)),
and a Word one that allows body-text rewriting while locking everything else
([word-body-rewrite.policy.yaml](examples/word-body-rewrite.policy.yaml)).
They are commented throughout and are the fastest way to see a real policy.

### Writing a policy: start from diff

You do not have to guess what to put in a policy. Make the edit you intend to
allow — by hand, or by letting the agent do it once — and ask what changed:

```bash
templategate diff --baseline plan.xlsx --candidate plan.edited.xlsx
```

Every change is listed with the exact location and attribute name a policy
rule would use. Allow the ones that were the point of the edit, and leave
everything else to default deny. Running `diff` on a *legitimate* edit is also
how you find out what your editing tool damages on the way past.

### Editing in real Excel

If the edit is made in Excel itself — by a person, or by an agent driving it
through COM — then a *correct, allowed* edit still changes more than the cell
you touched. Excel recalculates dependent formulas and refreshes chart caches
on every save, so the file legitimately differs in places nobody edited. A
policy that allows only the target cells will fail a perfectly good edit.

Editing a *single* quantity cell in a real workbook produced five changes: the
value itself, the cached results of two formulas that depend on it, the result
of an array formula, and the chart's cached copy of the series data.

The recipe:

```yaml
allow:
  - selector: "Sales!B3:B5"        # the cells actually being edited
    attributes: [value]
  - selector: "Sales!D3:D7"        # results of formulas that depend on them
    attributes: [value]
  - selector: "Sales!G3:G5"        # array-formula results — easy to forget
    attributes: [value]
  - selector: "package#charts:*"   # chart caches Excel refreshes on save
    attributes: [charts]
protect:
  - selector: "*"
    attributes: [formula, format, merge, conditional_formatting,
                 data_validation, print_settings, header_footer, vba]
```

Allowing `value` generously on computed ranges is safe **because `formula`
stays protected**: the results may move, but the rules that produce them may
not. A formula overwritten with a literal is still caught, under `formula`.

Allowing `package#charts:*` does mean a deliberate edit to a chart's own
definition would pass. The chart's *presence* is still guarded — deleting it
also removes the drawing that anchors it to the sheet, which `drawings`
catches.

The cascade is not really about Excel. It follows from **the baseline and the
candidate having been written by different tools**, and it happens in both
directions.

Run it the other way — a library editing a workbook that Excel authored, which
is the common case when an agent touches a real business document — and you
get the mirror image. openpyxl does not recalculate, so rather than refreshing
the cached results it *discards* them: every dependent cell reads as empty
until Excel next opens the file. One measured edit of three quantity cells
produced sixteen changes: the edit itself, seven wiped formula results, a date
cell whose value and number format were both rewritten, two hidden columns
whose widths read back differently, and a workbook extension block that was
dropped.

The fix is the same recipe — allow `value` across the computed ranges, keep
`formula` protected — and it is safe for the same reason: the cached answers
may move or vanish, the formulas that produce them may not. Expect to allow a
little more than values alone, such as `format` on a date cell the library
cannot round-trip and `layout` on hidden columns.

Steady state is quiet. When the baseline and the candidate come from the same
tool, there is no cascade and a narrow policy is exactly right.

### Editing in real Word

Word is well behaved in the steady state: opening a Word-authored document and
saving it without editing reports nothing at all. The friction is only on the
*first* save after Word takes over a file that a library produced.

That first save is a one-time re-authoring of the document's scaffolding
rather than damage. Word creates the footnote and endnote parts it expects
every document to have, drops legacy parts it no longer writes, rewrites the
style, numbering, settings and web-settings parts — re-pointing built-in
styles to its own identifiers — restamps the image markup in the body, and
re-points headers and footers at the renamed styles.

Those are real changes to real parts, which is why they are reported instead
of quietly ignored: only a person can say whether a wholesale style rewrite is
acceptable. In one measured run the first save produced ten changes across
those categories. Do not treat that as a fixed list — the built-in style names
Word writes are localized, so the exact set varies with version and locale.

**Re-baseline.** Save the document once in Word, adopt *that file* as the
baseline, and from then on baseline and candidate share an author and the
steady state is exactly zero.

One detail decides whether this works: adopt the file Word actually wrote. Do
not regenerate the baseline by converting the original a second time. Each
conversion stamps fresh identifiers, so two independent first saves of the
same source differ from each other even though neither was edited.

If you cannot re-baseline — the baseline is a fixed template someone else owns
— allow the known parts explicitly instead. These are the mechanical ones,
safe to allow outright:

```yaml
allow:
  - selector: "package#parts:word/footnotes.xml"
    attributes: [parts]
  - selector: "package#parts:word/endnotes.xml"
    attributes: [parts]
  - selector: "package#parts:word/numbering.xml"
    attributes: [parts]
  - selector: "package#parts:word/settings.xml"
    attributes: [parts]
  - selector: "package#parts:word/webSettings.xml"
    attributes: [parts]
```

That is deliberately not the whole residual — on the run measured above it
clears five of the ten changes. What remains is the style parts, the headers
and footers, and the image markup, and each of those needs a decision from
you, because each is somewhere real damage could hide. Derive the list from
your own `templategate diff` output rather than copying this one, since it is
specific to the install that produced the file.

Know what finishing the list costs. Adding `word/styles.xml` to it gives up
detection of style-definition tampering: setting every paragraph to 4pt white
text through a style change touches nothing *but* that part, so an allow rule
for it lets the change through silently. Re-baselining gives up nothing, which
is why it is the recommendation and this is the fallback.

### Choosing a mode

```yaml
mode: normal_input    # review_only | normal_input | page_extension
```

- `normal_input` (the default) compares by position: paragraph 12 is expected
  to still be paragraph 12. Every violation is an error.
- `page_extension` is for Word documents that are meant to grow. Paragraphs
  are aligned by content, so inserting one does not report every paragraph
  after it as changed, and a block that genuinely moved is reported as
  `moved`. On the same one-paragraph deletion, `normal_input` reports ten
  violations and `page_extension` reports one.
- `review_only` reports violations as warnings and exits 0. Useful for seeing
  what a policy would catch before you enforce it. It cannot bless a document
  that will not open.

### What a policy can name

`templategate init` writes a starter policy for the target you name, with the
attributes and structural keys already filled in — the fastest way to see the
current set. The ones worth knowing about:

- **Package parts** — `charts`, `comments`, `embedded`, `custom_xml`, `parts`
  and `links` apply to both formats; `pivot_tables` and `drawings` (shapes,
  textboxes and chart frames) are Excel-side, since a Word document keeps its
  equivalents in the document body where the Word attributes below cover them.
  `parts` is the catch-all for everything else in the file, including parts
  TemplateGate has never heard of, so losing one is damage by default. `links`
  compares the external targets of relationships, which is what catches a
  hyperlink repointed to a new URL while its display text is untouched.
- **Excel** — `layout` (hidden rows and columns, and their sizes),
  `protection` (sheet and workbook locking), `sheet_settings` (a workbook
  switched to manual calculation, or Excel's warning triangles suppressed).
- **Word** — `paragraph_format`, `field`, `bookmark`, `content_control`,
  `revision` (tracked changes), and `moved` for a block whose content survived
  but whose position did not. `markup` is the backstop: anything inside a
  block that no other attribute accounts for, which is what catches an
  unmodelled character style, a deleted footnote reference, a removed comment
  anchor, or a form field quietly disabled.

The VBA project keeps its own `vba` selector and attribute rather than living
under `package#`.

Locations are addressable:

```yaml
protect:
  - selector: "package#*"                              # every package part
  - selector: "package#charts:*"                       # one whole category
  - selector: "package#charts:xl/charts/chart1.xml"    # one specific part
  - selector: "package#links:https://example.com/x"    # one external target
  - selector: "Sheet1!1:10"                            # rows, for layout
  - selector: "'Q1!Q4'!A1"                             # see below
```

Word content controls and text boxes are `sdt1`, `textbox1`; a table nested
inside a table cell extends its parent's location, so the first cell of a
table nested in cell `r1c1` is `table1!r1c1!table1!r1c1`.

Sheet names are quoted Excel-style when they contain `'`, `!` or `#`:
`'Q1!Q4'!A1` is cell A1 of the sheet named `Q1!Q4`, where bare `Q1!Q4` would
mean cell Q4 of a sheet named `Q1`. You only need the quotes for names
containing those characters.

### Optional semantic checks

Everything above is deterministic and offline. The `semantic` block adds
judgement calls that structure alone cannot make — "the dates must not
contradict each other" — by handing the baseline and candidate text to a
command you choose:

```yaml
semantic:
  mode: "off"           # off (default) | review | gate
  provider: command
  command: "claude -p"  # any CLI that reads a prompt on stdin, prints JSON
  model: ""             # passed to that command as $TEMPLATEGATE_MODEL
  checks:
    - "dates and periods must not contradict each other"
```

`mode: off` is the default and makes no network calls at all. `review`
reports findings as warnings that never change PASS/FAIL; `gate` lets a
failed finding fail the run. The command receives a prompt on stdin and must
print a JSON array of `{check, verdict, message}`, where verdict is
`pass`, `fail` or `warning` — so any model or vendor works, and TemplateGate
pins none. With a non-`off` mode and no command set, the run reports a
configuration error rather than silently skipping the checks.

### Reading a report

Reports name what actually changed: a format violation reads
`font.bold True -> False`, not "the format changed". When one edit shifts
everything after it, the text and Markdown reports collapse the knock-on
changes into a single line — `p4..p11: content shifted because 1 paragraph
removed at p3` — while the JSON report always keeps every individual change,
tied together by a `group` field.

Other commands:

```bash
templategate diff --baseline a.xlsx --candidate b.xlsx   # list every change, no policy
templategate snapshot file.docx                           # dump the structural snapshot
```

## Which editing tools work

All of them, for detection. The gate judges the file, not the editor, so no
tool is unsupported. What differs is how much collateral each one leaves
behind, and therefore how wide your policy has to be to accept a correct edit.

Measured 2026-07-31 by opening one Excel-authored workbook — formulas with
cached results, hidden columns, a seal image, a Japanese date format, theme
colours — and saving it again with **no edit at all**:

| Writer | Version | Changes | What it did |
|---|---|---|---|
| Excel / Word via COM (incl. xlwings) | Office 16 | 0 | Word→Word is exactly silent; Excel refreshes caches it can recompute |
| python-docx | 1.2.0 | 0 | silent on a Word-authored document |
| openpyxl | 3.1.5 | 13 | discards 11 cached formula results; hidden-column width; workbook extension block |
| ExcelJS | 4.4.0 | 17 | vertical alignment and theme font colours on 12 cells; Japanese date format down to a bare serial |
| SheetJS (`xlsx`) | 0.18.5 | 49 | borders, fills, fonts and alignment across 38 cells; the image; the theme and metadata parts |
| pandas + XlsxWriter | 3.0.5 / 3.2.9 | 113 | authoring, not editing — see below |

Two caveats on that table. **openpyxl needs Pillow**: without it, it also
deletes images outright rather than preserving them, so install it alongside.
And **python-docx's zero is for this document** — earlier sweeps found it drops
comments and VBA, which this one did not contain.

**pandas and XlsxWriter are authoring tools, not editors.** XlsxWriter cannot
open an existing workbook at all; writing "to" a file replaces it. The 113
changes above are not a bug, they are what "rewrite the file from a DataFrame"
means — formulas, images, merges and formatting are gone because they were
never read. Re-baselining from the writer's own first output is the expected
workflow here, not a workaround.

### Getting out of the way of it

This kind of damage is deterministic: the same library on the same file makes
the same omissions every time. Retrying is pointless. Three things do work:

- **Use a richer editor.** A document with charts, shapes and real formatting
  needs a tool that understands them. Driving real Excel or Word through COM —
  xlwings, or the automation interface directly — is the safest option and is
  available on any machine with Office installed.
- **Edit the package directly.** For a value-only change, rewriting the one
  XML part inside the `.xlsx` zip touches nothing else by construction.
- **Let real Office repair what it can.** Opening a library-damaged workbook in
  Excel and saving restores the cached results it can recompute. In the run
  above that took openpyxl's 14 changes down to 2 — but the two that remained
  were the deleted image and the dropped extension block. **A resave recomputes;
  it does not resurrect.** Anything actually removed is still gone.

## When the gate fails

TemplateGate never writes to your documents. There is no repair mode and there
will not be one — a tool that can edit the file it is judging is a tool you
have to trust twice. Recovery is deliberately outside the tool, and you
already hold everything it needs:

**The baseline is the backup.** It is the untouched original, and a failed
candidate is a scratch file. The loop:

1. Read the violations. They name what broke and where.
2. Throw the candidate away. Do not patch it, and above all do not re-save it
   — whatever dropped a part on the first save will drop more on the second.
3. Copy the baseline again and retry, handing the violation report back to the
   agent as feedback. The JSON report is built to be fed straight in.
4. If the same violation survives two attempts, the tool doing the editing
   cannot preserve that part. Change the approach, not the policy.

In CI the baseline lives in git history, so every commit is another generation
of backup and you can always compare against the last version that passed.

Never widen the policy to make a failing check pass. That turns a caught error
into a silent one, which is the exact failure this tool exists to prevent.

## Python API

```python
import templategate

result = templategate.check("baseline.xlsx", "candidate.xlsx", "policy.yaml")
if not result.passed:
    for v in result.violations:
        print(v.change.location, v.change.attribute, v.message)
```

## For AI agents

The `skills/office-document-regression/` directory contains an agent skill
(Claude Code, Codex, ChatGPT and compatible) that teaches an agent the safe
workflow: edit a copy, run `templategate check`, interpret the JSON report,
and **never** edit the policy to make a failing check pass.

## GitHub Action

`action.yml` lives in the `action/` subdirectory, so the `uses:` path has to
include it. The runner fetches TemplateGate itself — you only check out *your*
repository so the documents and the policy are on disk.

```yaml
- uses: actions/checkout@v4
- uses: actions/setup-python@v5
  with:
    python-version: "3.12"
- uses: kousunh/TemplateGate/action@v1
  with:
    baseline: docs/plan_baseline.xlsx
    candidate: docs/plan.xlsx
    policy: .templategate/plan.policy.yaml
```

The step fails the job when the check fails, and appends the report to the job
summary either way. It exposes two outputs, `passed` (`true` / `false`) and
`report-path`; to read them instead of failing the job, add
`continue-on-error: true` to the step.

## What TemplateGate is not

- Not an editor, converter, or auto-repair tool. See
  [When the gate fails](#when-the-gate-fails).
- Changes inside a part that is not separately modelled are reported as that
  part changing, without naming the property — you learn `word/header1.xml`
  was modified, not which run turned white.
- No round-trip normalization (differences introduced by re-saving in another
  Office application are out of scope).
- No visual/PDF regression.
- Semantic mode `off` performs zero network calls; TemplateGate itself never
  uploads your documents anywhere.

## License

MIT. See [LICENSE](LICENSE).

TemplateGate is an independent open-source project and is not affiliated with,
endorsed by, or sponsored by Microsoft. "Microsoft", "Office", "Excel" and
"Word" are trademarks of Microsoft Corporation, used here only to identify
file formats.
