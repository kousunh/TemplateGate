# TemplateGate

**Policy-as-code acceptance gate for AI-edited Excel & Word documents.**

日本語版は [README.ja.md](README.ja.md) をどうぞ。

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

## Install

```bash
pip install templategate
```

Requires Python 3.10+. Pure-Python dependencies only (openpyxl, python-docx, PyYAML).

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
semantic:
  mode: "off"                     # off | review | gate
```

Other commands:

```bash
templategate diff --baseline a.xlsx --candidate b.xlsx   # list every change, no policy
templategate snapshot file.docx                           # dump the structural snapshot
```

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

```yaml
- uses: <owner>/templategate/action@v1
  with:
    baseline: docs/plan_baseline.xlsx
    candidate: docs/plan.xlsx
    policy: .templategate/plan.policy.yaml
```

## What TemplateGate is not

- Not an editor, converter, or auto-repair tool.
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
