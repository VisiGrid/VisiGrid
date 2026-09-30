#!/usr/bin/env python3
"""Independent, stdlib-only OOXML audit. No spreadsheet library rewrites input.

Usage: python3 scripts/xlsx-audit.py SOURCE_DIRECTORY OUTPUT_DIRECTORY
Build first: cargo build --locked -p visigrid-io --example xlsx_probe
This is an exploratory report, not a lossless compatibility certification.
"""
import decimal
import hashlib
import json
import pathlib
import posixpath
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def snapshot(path):
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        if bad:
            raise ValueError(f"bad ZIP CRC: {bad}")
        # Parse every XML part, including relationships that VisiGrid may ignore.
        for part in z.namelist():
            if part.endswith((".xml", ".rels")):
                ET.fromstring(z.read(part))
        def read(part):
            return ET.fromstring(z.read(part))
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ["".join(si.itertext()) for si in read("xl/sharedStrings.xml")]
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in read("xl/_rels/workbook.xml.rels")}
        result = {"sheets": [], "parts": {}}
        for kind in ("charts", "tables", "drawings", "pivotTables", "externalLinks"):
            result["parts"][kind] = sum(p.startswith(f"xl/{kind}/") and p.endswith(".xml") and "/_rels/" not in p for p in z.namelist())
        book = read("xl/workbook.xml")
        result["defined_names"] = [(n.attrib, n.text) for n in book.findall("s:definedNames/s:definedName", NS)]
        result["date1904"] = (book.find("s:workbookPr", NS).get("date1904", "0") if book.find("s:workbookPr", NS) is not None else "0")
        for sheet in book.findall("s:sheets/s:sheet", NS):
            target = rels[sheet.attrib[RID]]
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            root = read(part)
            cells = {}
            for c in root.findall("s:sheetData/s:row/s:c", NS):
                v = c.findtext("s:v", None, NS)
                f = c.find("s:f", NS)
                t = c.get("t", "n")
                if f is not None:
                    # Shared followers require expansion; keep explicit metadata
                    # so this report cannot silently claim formula equivalence.
                    value = ["formula", f.text or "", dict(f.attrib)]
                elif t == "s" and v is not None:
                    value = ["text", strings[int(v)]]
                elif t == "inlineStr":
                    value = ["text", "".join(c.find("s:is", NS).itertext())]
                elif v is None:
                    continue
                elif t == "n":
                    value = ["number", str(decimal.Decimal(v).normalize())]
                else:
                    value = [t, v]
                cells[c.attrib["r"]] = value
            pane = root.find("s:sheetViews/s:sheetView/s:pane", NS)
            result["sheets"].append({
                "name": sheet.get("name"), "state": sheet.get("state", "visible"), "cells": cells,
                "merges": sorted(m.get("ref") for m in root.findall("s:mergeCells/s:mergeCell", NS)),
                "conditional_rules": len(root.findall("s:conditionalFormatting/s:cfRule", NS)),
                "validations": len(root.findall("s:dataValidations/s:dataValidation", NS)),
                "frozen": {k: pane.get(k, "0") for k in ("xSplit", "ySplit")} if pane is not None and pane.get("state", "").startswith("frozen") else {},
            })
        return result


def compare(before, after, edited=False):
    diffs = []
    for key in ("parts", "defined_names", "date1904"):
        if before[key] != after[key]:
            diffs.append({"feature": key, "before": before[key], "after": after[key]})
    if [s["name"] for s in before["sheets"]] != [s["name"] for s in after["sheets"]]:
        diffs.append({"feature": "sheet_order_or_names"})
    for i, (a, b) in enumerate(zip(before["sheets"], after["sheets"])):
        for key in ("state", "merges", "conditional_rules", "validations", "frozen"):
            if a[key] != b[key]:
                diffs.append({"sheet": a["name"], "feature": key, "before": a[key], "after": b[key]})
        for cell in sorted(a["cells"].keys() | b["cells"].keys()):
            old = ["text", "VisiGrid QA edit"] if edited and i == 0 and cell == "A1" else a["cells"].get(cell)
            new = b["cells"].get(cell)
            if old != new:
                diffs.append({"sheet": a["name"], "cell": cell, "feature": "cell", "before": old, "after": new})
    return diffs


def check_basics(path):
    """Check known intended formatting, independently of VisiGrid's reader.

    Only the generated basics fixture has these expectations. This does not
    pretend to normalize every possible Excel style in the external corpus.
    """
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        styles = ET.fromstring(z.read("xl/styles.xml"))
        cells = {c.get("r"): c for c in root.findall("s:sheetData/s:row/s:c", NS)}
        xfs = styles.find("s:cellXfs", NS)
        def xf(cell):
            return xfs[int(cells[cell].get("s", "0"))]
        def component(cell, kind, attr):
            return styles.find("s:" + kind, NS)[int(xf(cell).get(attr, "0"))]
        def fill(cell):
            return component(cell, "fills", "fillId").find("s:patternFill/s:fgColor", NS).get("rgb")
        font = component("A1", "fonts", "fontId")
        assert font.find("s:b", NS) is not None, "title bold lost"
        assert font.find("s:color", NS).get("rgb") == "FFFFFFFF", "title white lost"
        assert fill("A1") == "FF245A81", "title blue lost"
        assert fill("A11") == "FFE2F0D9", "merge green lost"
        border = component("A11", "borders", "borderId")
        for side in ("left", "right", "top", "bottom"):
            assert border.find("s:" + side, NS).get("style") == "thin", f"{side} border lost"
        fmts = {n.get("numFmtId"): n.get("formatCode") for n in styles.findall("s:numFmts/s:numFmt", NS)}
        for cell, code in {"B3": "$#,##0.00", "B4": "yyyy-mm-dd", "B5": "0.00%"}.items():
            assert fmts[xf(cell).get("numFmtId")] == code, f"{cell} number format lost"
        rows = {r.get("r"): r for r in root.findall("s:sheetData/s:row", NS)}
        assert float(rows["1"].get("ht")) == 28, "title height lost"
        assert rows["13"].get("hidden") == "1", "hidden row lost"
        cols = root.findall("s:cols/s:col", NS)
        for index, width in ((1, 30.7109375), (2, 24.7109375)):
            col = next(c for c in cols if int(c.get("min")) <= index <= int(c.get("max")))
            assert abs(float(col.get("width")) - width) < 0.01, "column width drift"
        assert any(int(c.get("min")) <= 5 <= int(c.get("max")) and c.get("hidden") == "1" for c in cols), "hidden column lost"
        return "passed: RGB colors, bold, borders, number formats, sizes, hidden row/column"


def main():
    source, output = map(pathlib.Path, sys.argv[1:])
    source, output = source.resolve(), output.resolve()
    if source == output or source in output.parents:
        raise ValueError("output must be outside the source directory")
    root = pathlib.Path(__file__).resolve().parents[1]
    exe = root / "target/debug/examples" / ("xlsx_probe.exe" if sys.platform == "win32" else "xlsx_probe")
    output.mkdir(parents=True, exist_ok=True)
    manifest = json.loads((root / "fixtures/xlsx-corpus.json").read_text())
    checksums = {f["file"]: f["sha256"] for f in manifest["files"]}
    report = []
    failures = 0
    for path in sorted(source.glob("*.xlsx")):
        dest = output / path.stem
        item = {"file": path.name, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
        try:
            if path.name in checksums and item["sha256"] != checksums[path.name]:
                raise ValueError("fixture checksum differs from pinned manifest")
            original = snapshot(path)
            run = subprocess.run([str(exe), str(path), str(dest)], capture_output=True, text=True, timeout=90)
            (output / (path.stem + ".log")).write_text(run.stdout + run.stderr)
            if run.returncode:
                raise RuntimeError(f"probe exit {run.returncode}; see log")
            item["sheets"] = len(original["sheets"])
            item["cells"] = sum(len(s["cells"]) for s in original["sheets"])
            item["paths"] = {}
            for name in ("unchanged", "edited", "native-roundtrip"):
                item["paths"][name] = compare(original, snapshot(dest / (name + ".xlsx")), name == "edited")
            if path.name == "basics.xlsx":
                item["basic_format_checks"] = {name: check_basics(dest / (name + ".xlsx")) for name in item["paths"]}
            item["native_vs_direct"] = compare(snapshot(dest / "unchanged.xlsx"), snapshot(dest / "native-roundtrip.xlsx"))
            print(path.name, {k: len(v) for k, v in item["paths"].items()}, flush=True)
        except Exception as e:
            item["error"] = str(e)
            failures += 1
            print(path.name, str(e), flush=True)
        report.append(item)
    (output / "report.json").write_text(json.dumps(report, indent=2))
    print(f"Report: {output / 'report.json'}")
    # Differences remain observations until reviewed expectations are assigned.
    return 1 if failures or not report else 0


if __name__ == "__main__":
    sys.exit(main())
