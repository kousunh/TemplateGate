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
