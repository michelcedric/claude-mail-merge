#!/usr/bin/env python3
"""Move generated PDFs from WORK to OUTPUT and write an XML manifest.

  publish.py --report work/merge_report.json --output output [--manifest manifest.xml] [--clean]
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import shutil
import sys
import xml.etree.ElementTree as ET


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--report", default="work/merge_report.json")
    ap.add_argument("--output", default="output")
    ap.add_argument("--manifest", default="manifest.xml")
    ap.add_argument("--clean", action="store_true", help="delete the work folder afterwards")
    args = ap.parse_args()

    if not os.path.exists(args.report):
        print(json.dumps({"ok": False, "error": f"report not found: {args.report}"}))
        sys.exit(1)
    with open(args.report, encoding="utf-8") as fh:
        report = json.load(fh)
    work = os.path.dirname(os.path.abspath(args.report))
    os.makedirs(args.output, exist_ok=True)

    root = ET.Element("manifest", {
        "generated": dt.datetime.now().isoformat(timespec="seconds"),
        "source": report["excel"],
        "sheet": report["sheet"],
        "template": report["template"],
    })
    files_el = ET.SubElement(root, "files")
    moved, missing, overwritten = [], [], []

    for item in report["generated"]:
        src = os.path.join(work, item["file"])
        dst = os.path.join(args.output, item["file"])
        if not os.path.exists(src):
            missing.append(item["file"])
            continue
        if os.path.exists(dst):
            overwritten.append(item["file"])
        shutil.move(src, dst)
        f = ET.SubElement(files_el, "file", {
            "name": item["file"],
            "row": str(item["row"]),
            "size": str(os.path.getsize(dst)),
            "sha256": sha256(dst),
        })
        data = ET.SubElement(f, "data")
        for k, v in item["data"].items():
            ET.SubElement(data, "field", {"name": k}).text = v
        moved.append(item["file"])

    root.set("count", str(len(moved)))
    if report.get("errors"):
        errs = ET.SubElement(root, "errors")
        for e in report["errors"]:
            ET.SubElement(errs, "error", {"row": str(e["row"])}).text = e["error"]

    ET.indent(root) if hasattr(ET, "indent") else None
    manifest_path = os.path.join(args.output, args.manifest)
    ET.ElementTree(root).write(manifest_path, encoding="utf-8", xml_declaration=True)

    if args.clean and not missing:
        shutil.rmtree(work, ignore_errors=True)

    print(json.dumps({
        "ok": not missing,
        "output": os.path.abspath(args.output),
        "manifest": manifest_path,
        "published": len(moved),
        "overwritten": overwritten,
        "missing": missing,
        "merge_errors": len(report.get("errors", [])),
    }, indent=2, ensure_ascii=False))
    sys.exit(0 if not missing else 1)


if __name__ == "__main__":
    main()
