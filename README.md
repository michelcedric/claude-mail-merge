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

## Visual Studio Mail Merge agent

The repository includes a selectable **Mail Merge** custom agent at `.github/agents/mail-merge.agent.md`.

1. Open this repository in Visual Studio and open **GitHub Copilot Chat**.
2. Use the agent picker in Copilot Chat to select **Mail Merge**. If it is not listed, reload the repository or Visual Studio so the `.github/agents/` profile is discovered.
3. Place exactly one `.xlsx` or `.xlsm` workbook and one `.docx` template in the repository's `input/` folder. The first spreadsheet row must contain the merge-field headers, and those names must match the template fields exactly.
4. Ask the agent to run the mail merge. It checks the input before proceeding, uses the repository's merge and publish scripts, and asks before file overwrites or moves/deletions. Review those prompts and confirm only if you want the operation to continue.
5. The merge report and intermediate PDFs are kept in `work/` until publishing is confirmed; published PDFs and `manifest.xml` are written to `output/`.

If the agent reports that no PDF converter is available, install LibreOffice and make its `soffice` executable available on `PATH`, then run the agent again. LibreOffice is recommended for faithful PDF output; Chrome conversion is intended only for simple layouts.
