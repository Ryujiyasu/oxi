"""Fuzz Quirk Discovery — measurement pipeline.

For each fuzz docx in a batch:
1. Word COM: open, extract per-paragraph (and per-cell) y positions via
   Information(6) on collapsed start range
2. Oxi GDI: run oxi-gdi-renderer --dump-layout, extract text positions

Output per-doc JSON pairing Word vs Oxi positions.
"""
from __future__ import annotations
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(__file__).parent.parent.parent
RENDERER = ROOT / "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"


def collapse_y(rng):
    doc = rng.Document
    return doc.Range(rng.Start, rng.Start).Information(6)


def measure_word(word, docx_path: Path) -> dict:
    doc = word.Documents.Open(str(docx_path.absolute()), ReadOnly=True)
    try:
        result = {
            "head_y": None,
            "anchor_y": None,
            "rows": [],
        }
        # body paras
        for p in doc.Paragraphs:
            t = p.Range.Text.strip()
            if t == "HEAD":
                result["head_y"] = collapse_y(p.Range)
            elif t == "ANCHOR":
                result["anchor_y"] = collapse_y(p.Range)
        if doc.Tables.Count > 0:
            tbl = doc.Tables(1)
            for r in range(1, tbl.Rows.Count + 1):
                row = tbl.Rows(r)
                row_data = {"row_idx": r - 1, "cells": []}
                for c in range(1, row.Cells.Count + 1):
                    cell = row.Cells(c)
                    cell_data = {"col_idx": c - 1, "y": collapse_y(cell.Range)}
                    row_data["cells"].append(cell_data)
                result["rows"].append(row_data)
        return result
    finally:
        doc.Close(SaveChanges=False)


def measure_oxi(docx_path: Path) -> dict:
    with tempfile.TemporaryDirectory(prefix="fuzz_oxi_") as tmp:
        prefix = os.path.join(tmp, "p_")
        dump = os.path.join(tmp, "layout.json")
        proc = subprocess.run(
            [str(RENDERER), str(docx_path), prefix, "--dump-layout=" + dump],
            capture_output=True, text=True, timeout=60,
        )
        if proc.returncode != 0:
            return {"error": f"renderer rc={proc.returncode}: {proc.stderr[:300]}"}
        try:
            with open(dump, encoding="utf-8") as f:
                d = json.load(f)
        except Exception as e:
            return {"error": str(e)}

    result = {"head_y": None, "anchor_y": None, "rows": {}}
    pages = d.get("pages", [])
    if not pages:
        return {"error": "no pages"}
    p0 = pages[0]
    for el in p0.get("elements", []):
        if el.get("type") != "text":
            continue
        cr = el.get("cell_row_idx")
        cc = el.get("cell_col_idx")
        if cr is None:
            text = el.get("text", "")
            if "HEAD" in text:
                result["head_y"] = el["y"]
            elif "ANCHOR" in text:
                result["anchor_y"] = el["y"]
        else:
            key = (cr, cc)
            if key not in result["rows"] or el["y"] < result["rows"][key]["y"]:
                result["rows"][key] = {"y": el["y"], "x": el["x"]}
    # convert to list
    rows_list = []
    for (cr, cc), v in sorted(result["rows"].items()):
        rows_list.append({"row_idx": cr, "col_idx": cc, "y": v["y"], "x": v["x"]})
    result["rows"] = rows_list
    return result


def main(batch_name: str = "smoke"):
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    batch_dir = Path(__file__).parent / "fuzz_runs" / batch_name
    if not batch_dir.exists():
        print(f"Batch dir not found: {batch_dir}")
        return

    docx_files = sorted(batch_dir.glob("fuzz_*.docx"))
    print(f"Measuring {len(docx_files)} docs from {batch_dir}")
    results = []
    try:
        for p in docx_files:
            print(f"  {p.name}", end=" ... ", flush=True)
            try:
                w = measure_word(word, p)
                o = measure_oxi(p)
                results.append({"doc": p.name, "word": w, "oxi": o})
                print("OK")
            except Exception as e:
                results.append({"doc": p.name, "error": str(e)})
                print(f"ERR {e}")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    out = batch_dir / "measurements.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"Wrote {out}")


if __name__ == "__main__":
    batch = sys.argv[1] if len(sys.argv) > 1 else "smoke"
    main(batch_name=batch)
