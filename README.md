# claude-mail-merge

A Claude Code mail-merge pipeline: one Excel file plus one Word template with MERGEFIELDs become one PDF per row, published with an XML manifest.

## Layout

```
.claude/
  agents/
    merge-agent.md        # input/ -> PDFs in work/
    publish-agent.md      # work/ -> output/ + manifest.xml
  skills/
    run-mailmerge/        # orchestrator, invoked with /run-mailmerge
    mail-merge/           # merge script (scripts/merge.py)
    publish-output/       # publish script (scripts/publish.py)
requirements.txt
```

## Usage

1. Put exactly one `.xlsx` (first row = column headers) and one `.docx` template in `input/`.
   Template field names must match the Excel headers exactly.
2. In Claude Code, run `/run-mailmerge`, optionally with a file-name pattern such as
   `/run-mailmerge {InvoiceNo}_{LastName}`, plus `keep-work` or `allow-missing`.
3. The PDFs and `manifest.xml` land in `output/`.

The merge agent creates `.venv` and installs `requirements.txt` on first run.
LibreOffice gives the most faithful PDFs (`brew install --cask libreoffice`); Word and Chrome are fallbacks.

`input/`, `work/` and `output/` are git-ignored, so customer data and generated PDFs stay local.
