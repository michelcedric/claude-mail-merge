---
name: publish-output
description: Publish generated PDFs from the work folder to the output folder and write an XML manifest listing every generated file. Use after a mail merge has produced work/merge_report.json.
user-invocable: false
---

# Publish PDFs + XML manifest

Run from the project root:

```bash
.venv/bin/python .claude/skills/publish-output/scripts/publish.py \
  --report work/merge_report.json --output output --manifest manifest.xml --clean
```

- `--report`: the handoff file written by the mail-merge step. If it is missing, stop and report it. Do not try to rebuild it.
- `--clean`: deletes `work/` after a fully successful move. Leave it off if the caller asked to keep intermediate files.
- `--manifest`: manifest file name (default `manifest.xml`). Use the caller's name if given.

Existing files with the same name in `output/` are overwritten and listed in `overwritten`.

## Manifest format (`output/manifest.xml`)

```xml
<?xml version='1.0' encoding='utf-8'?>
<manifest generated="2026-10-23T10:00:00" source="data.xlsx" sheet="Sheet1" template="template.docx" count="3">
  <files>
    <file name="INV001_Dupont.pdf" row="2" size="48213" sha256="…">
      <data>
        <field name="InvoiceNo">INV001</field>
        …
      </data>
    </file>
  </files>
  <errors>                       <!-- only when merge rows failed -->
    <error row="5">…</error>
  </errors>
</manifest>
```

## Verify

- stdout JSON: `ok` is true, `published` equals the number of generated files, and `missing` is empty.
- `ls output/` shows the PDFs plus the manifest.
- Read `output/manifest.xml` and confirm `count` matches.
