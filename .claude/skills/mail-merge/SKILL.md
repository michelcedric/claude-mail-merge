---
name: mail-merge
description: Mail merge an Excel file into a Word template containing MERGEFIELDs and produce one PDF per Excel row. Use when asked to mail merge, generate letters/invoices/certificates from a spreadsheet and a .docx template.
user-invocable: false
---

# Mail merge (Excel + Word template → PDFs)

All commands run from the project root with the project virtualenv:

```bash
PY=.venv/bin/python
SCRIPT=.claude/skills/mail-merge/scripts/merge.py
```

If `.venv` does not exist, create it first:
`python3 -m venv .venv && .venv/bin/pip install -r requirements.txt`

## Inputs

- `input/` must contain exactly **one** `.xlsx` (first row = column headers) and **one** `.docx` template.
- The template uses Word merge fields (Insert → Quick Parts → Field → MergeField). Field names must match the Excel headers exactly (case-sensitive).

## Step 1 — Inspect (always do this first)

```bash
$PY $SCRIPT inspect --input input
```

Returns JSON with `columns`, `merge_fields`, `fields_missing_in_excel`, `row_count`, `sample_row`, and `converter`.

Check before going on:
- `fields_missing_in_excel` is empty. If not, **stop** and report the mismatch, unless the caller explicitly accepted blank fields (then add `--allow-missing`). Suggest the closest column name when it is an obvious typo or case difference.
- `row_count` > 0.
- `converter` is not null. Quality ranking: `libreoffice` (best) > `word` > `chrome` (fallback: simple layouts only, headers/footers and complex positioning are lost). When it is `chrome`, mention in the result that installing LibreOffice (`brew install --cask libreoffice`) gives faithful PDFs.

## Step 2 — Choose file names

`--name-pattern` uses Python format syntax with Excel column names plus `{row}` (Excel row number).
Pick a pattern from identifying columns visible in `columns` (e.g. an ID, number or name), for example `"{InvoiceNo}_{LastName}"`. Use the caller's pattern if given. Default: `document_{row}`.
Unsafe characters are replaced automatically, and duplicates get a `_2`, `_3` suffix.

## Step 3 — Run

```bash
$PY $SCRIPT run --input input --work work --name-pattern "{InvoiceNo}_{LastName}"
```

Options: `--sheet NAME` (default first sheet), `--date-format "%d/%m/%Y"`, `--allow-missing`, `--keep-docx` (keep the intermediate .docx for debugging).

Output:
- PDFs in `work/`
- `work/merge_report.json`: the handoff file for publishing (source files, converter, per-row `generated` list with file name and row data, `errors`)
- stdout JSON summary: `ok`, `report`, `rows`, `generated`, `errors`

## Step 4 — Verify

- `generated` == `rows` and `errors` is empty → success.
- Spot-check one PDF: `ls -la work/*.pdf`, then read one PDF to confirm the merged values appear (no leftover `«FieldName»`).
- Partial failure: keep going, but list every failed row with its error.

## Never

- Never edit the files in `input/`.
- Never write to `output/`. Publishing is a separate step (publish-output skill).
