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

## Hard rules

- **NEVER modify the policy file to make a failing check pass.** The policy
  is the user's trusted specification. If you believe the policy is wrong,
  propose a change and wait for explicit user approval.
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
