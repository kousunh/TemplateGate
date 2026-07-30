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
     `embedded`, or `custom_xml`. Read these as: **the library you edited
     with could not represent that part, so it was dropped on save.** You
     usually cannot repair this by editing the candidate further, because
     the content is simply gone. Start over from a fresh copy of the
     baseline and make the change in a way that preserves the whole file —
     and do not re-save the candidate again, which only compounds the loss.
   - exit code `2`: an execution error (bad path, invalid policy). Fix the
     invocation, not the document.

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
