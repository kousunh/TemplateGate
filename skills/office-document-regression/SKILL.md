---
name: office-document-regression
description: >
  Safely edit Excel (.xlsx/.xlsm) or Word (.docx) documents and verify the
  result with TemplateGate. Use whenever asked to modify an Office document
  that must keep its structure, formulas, or formatting intact — before
  handing the edited file back, always run the acceptance check.
---

# Office document regression check (TemplateGate)

You are editing an Office document whose structure must not break. Follow
this workflow exactly.

## Workflow

1. **Never edit the original.** Copy it first:
   - baseline: the untouched original file
   - candidate: the copy you edit
2. **Read the policy** (`*.policy.yaml`, usually provided by the user or the
   repository). It defines what you may change (`allow`) and what must never
   change (`protect`, `structural`). If no policy exists, ask the user for
   one or propose a draft **for the user to review** — see the hard rules.
3. **Make your edits on the candidate copy.**
4. **Run the check:**

   ```bash
   templategate check --baseline <original> --candidate <edited> \
     --policy <policy.yaml> --report json
   ```

5. **Interpret the result:**
   - exit code `0` (PASS): report the allowed changes to the user and deliver
     the candidate file.
   - exit code `1` (FAIL): read `violations[]` in the JSON. Each entry has
     `change.location` (e.g. `Sheet1!D8`), `change.attribute` (e.g.
     `formula`), and `message`. Fix your edit so the violation disappears —
     e.g. restore the destroyed formula — and re-run. If you cannot fix it,
     deliver nothing: quarantine the candidate, show the violations to the
     user, and let them decide.

     One broken cell often yields more than one entry: value and formula are
     compared independently, so overwriting `=SUM(D4:D6)` with a literal
     reports both a `value` violation and a `formula` one at the same
     location. Restoring the formula clears both — fix the cause, not each
     entry.

     A violation whose location starts with `package#` means part of the
     document itself went missing or changed — attribute `charts`,
     `pivot_tables`, `drawings` (shapes and textboxes), `comments`,
     `embedded`, `custom_xml`, `parts` (anything else in the file, including
     parts the tool does not model), or `links` (an external hyperlink
     target). Read these as: **the library you edited with could not
     represent that part, so it was dropped on save.** You usually cannot
     repair this by editing the candidate further, because the content is
     simply gone. Start over from a fresh copy of the baseline — and do not
     re-save the candidate, which only compounds the loss.

     A FAIL can also mean the candidate is *damaged*: it still opens as a
     package but has lost parts it referenced. That is a failed check, not a
     tool error, and it is not repairable by further editing.

     A `markup` violation means something structural changed inside that
     block that is not its text and not its formatting — a footnote
     reference or comment anchor that got dropped, a form field disabled, a
     character style removed. Retyping the text will not fix it, because the
     text was never the problem. Restart from the baseline.
   - exit code `2`: an execution error (bad path, invalid policy, or a file
     that will not open at all). Fix the invocation, not the document.

## The retry loop

TemplateGate never writes to the document, so it cannot repair anything. The
baseline is the backup, and a failed candidate is a scratch file:

1. Read `violations[]`.
2. Delete the candidate. Do not patch it and do not re-save it.
3. Copy the baseline again, redo the edit, and re-run the check — using the
   previous report as your own feedback about what to avoid this time.
4. If the same violation appears twice, the method you are using cannot
   preserve that part of the file. Change the method, tell the user, and stop.
   Do not change the policy.

## What the check can see

Assume everything is visible. Beyond cell values and paragraph text, the check
covers hidden rows and columns, sheet protection, manual-calculation mode,
hyperlink targets even when the display text is unchanged, paragraph
formatting, fields, bookmarks, content controls, text boxes (`sdt<N>` /
`textbox<N>`), nested tables, and every package part in the file.

Editing through tracked changes hides nothing: a tracked insertion is text on
the page, so it is reported as a `text` change and a `revision` change.
Anything in a block that no named attribute covers is still compared, under
`markup`. There is no way to make an edit that the report will not mention.

## If you are driving Excel itself

Whenever the baseline and the candidate come from different tools, cells you
never touched will differ — and that is expected, not damage. Excel
recalculates dependent formulas and refreshes chart caches on every save; a
library does the opposite and discards the cached results, leaving those cells
empty until Excel next opens the file. So editing an Excel-authored workbook
with openpyxl cascades just as much as editing through Excel does.

If the policy allows only the cells you edited, a correct edit still FAILs.
Report that to the user as a policy needing widening on the computed ranges,
rather than assuming you broke the file. Do not edit the policy yourself.

## If the document is opened in real Word

The first time real Word saves a document that a library produced, it rewrites
the scaffolding — style, numbering and settings parts, footnote and endnote
parts it adds, legacy parts it drops. That is a one-time re-authoring, not
damage you caused, and it looks alarming because it is many parts at once.
Word-to-Word saves after that report nothing.

If you see it, say so and stop: the fix is for the user to re-baseline (save
once in Word and adopt that file as the baseline), which is their call, not
yours. Do not widen the policy to absorb it.

## When it FAILs, classify before you react

The first question is never "how do I fix this edit" — it is **whose limit did
I just hit**.

**Your mistake.** `text` or `value` violations at locations you deliberately
edited. You changed something you were not allowed to change. Fix the edit and
re-run. This is the only class that retrying helps.

**Your tool's ceiling.** `package#...`, `charts`, `images`, `parts`, `format`
or metadata violations you did not intend, usually far from where you were
working. The library you are using cannot represent that part, so it dropped
it on save. **This is deterministic — running the same library again produces
the same loss.** Retrying is not a strategy. Escalate instead:

1. Re-do the edit with real Excel or Word driven through COM (xlwings, or the
   automation interface). It is the richest editor available and is on any
   machine with Office installed.
2. If only a value needs changing, edit the XML part inside the package
   directly — by construction that touches nothing else.
3. If neither is available, stop and hand the violation list to the human.
   Say which tool you used and what it dropped. That is a useful report, not
   a failure to finish.

Do not try to repair a dropped part by editing the candidate further, and do
not re-save it — the second save loses more than the first.

## Hard rules

- **NEVER modify the policy file to make a failing check pass.** The policy
  is the user's trusted specification. If you believe the policy is wrong,
  propose a change and wait for explicit user approval.
- **NEVER run `templategate suggest` to widen the policy governing your own
  edit.** `suggest` drafts a policy from an edit a *human* has reviewed and
  approved. Pointing it at your own candidate and adopting the result would
  make you the author of the rules you are judged by, which defeats the whole
  arrangement — the check would then only confirm that you did what you did.
  If the user asks you to run it, draft from an edit they have approved,
  write it to a file, and leave adopting it to them.
- **NEVER deliver a candidate that fails the check** without telling the user
  it failed and why.
- Do not "fix" unexpected diffs by re-saving the baseline or swapping files.
- If the check reports formula/format/VBA violations you did not intend, or
  any `package#` violation, assume your editing tool corrupted the file —
  start over from a fresh copy.

## Useful commands

```bash
templategate diff --baseline a.xlsx --candidate b.xlsx --json  # inspect all changes
templategate snapshot file.xlsx                                 # what TemplateGate sees
templategate init --target excel                                # policy template (draft for the user)
```

When you show the outcome to the user, speak their language: `--report text
--lang ja` (or `en`) renders the human-readable report. Keep `--report json`
for your own loop — its structure never changes with the language.
