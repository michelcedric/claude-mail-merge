---
name: run-mailmerge
description: Orchestrate the full mail-merge pipeline. merge-agent turns input/ (Excel + Word template) into PDFs, then publish-agent moves them to output/ with an XML manifest. Use when the user asks to run the mail merge, generate the PDFs, or process the input folder.
argument-hint: "[name pattern, e.g. {InvoiceNo}_{LastName}] [keep-work] [allow-missing]"
disable-model-invocation: false
---

# Mail-merge orchestrator

You coordinate two subagents. Do **not** run the merge or publish scripts yourself: delegate with the Agent tool and act on the JSON they return.

Arguments: `$ARGUMENTS`. Any `{...}` text is the file-name pattern. `keep-work` means do not clean `work/`. `allow-missing` means blank fields are accepted.

## 1. Pre-flight (yourself, quick)
Run this exact command, so that a missing `work/` (the normal case after a clean run) does not end with exit code 1 and show up as a failed step:

```bash
ls -la input/; echo "--- work:"; ls -la work/ 2>/dev/null || echo "work/ is clean"
```

- `input/` must hold exactly one `.xlsx` and one `.docx` (ignore `~$` lock files). If not, tell the user what is wrong and stop.
- If `work/` contains leftovers from an earlier run, delete `work/` before starting.

Apply the same rule to any other check you run yourself, for example after publishing: a folder that is expected to be missing must not make the command fail (`2>/dev/null || echo "..."`).

## 2. Merge — delegate to `merge-agent`
Call the Agent tool with `subagent_type: "merge-agent"` and a prompt that includes:
- the file-name pattern (if given) and whether missing fields are allowed
- "Return only the JSON result block."

Handle the result:
- `failed` or `blocking_issue` set → **stop**. Explain the issue to the user in plain words, with the fix (for example "rename the Excel column `Last Name` to `LastName`"). Do not call publish-agent.
- `partial` → continue, but remember the failed rows for the final report.
- `success` → continue.

## 3. Publish — delegate to `publish-agent`
Call the Agent tool with `subagent_type: "publish-agent"` and a prompt containing the report path from step 2, plus `--clean` unless `keep-work` was requested. Ask for only the JSON result block.

If it fails, report the problem and say that the PDFs are still in `work/`.

## 4. Final report to the user
Write a short summary:
- the rows processed, the PDFs published and the converter used (when it is `chrome`, recommend LibreOffice for faithful layout)
- the output folder and manifest path
- any failed rows or overwritten files
- warnings returned by the agents
