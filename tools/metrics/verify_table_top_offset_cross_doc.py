"""Verify Oxi's table-top +2.5pt offset is universal or b35123-specific.

Test on additional table-heavy docs from bottom bucket: 2ea81a, 1ec1091, e3c545.
"""
import json
import sys
import time
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path("tools/golden-test/documents/docx")
OUT = Path("pipeline_data/table_top_offset_audit.json")

DOCS = [
    "b35123fe8efc_tokumei_08_01.docx",
    "2ea81a8441cc_0025006-192.docx",
    "1ec1091177b1_006.docx",
]


def measure(word, docx_path):
    last = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
            time.sleep(0.5)
            data = {"doc": docx_path.name, "tables": []}
            n_tbl = doc.Tables.Count
            for ti in range(1, min(n_tbl + 1, 5)):  # first 4 tables
                tbl = doc.Tables(ti)
                tbl_top = tbl.Range.Information(6)
                tbl_page = tbl.Range.Information(3)
                cells = []
                # First 2 cells of first 2 rows
                for ri in range(1, min(tbl.Rows.Count + 1, 3)):
                    for ci in range(1, min(tbl.Columns.Count + 1, 3)):
                        try:
                            cell = tbl.Cell(ri, ci)
                            first_char = doc.Range(cell.Range.Start, cell.Range.Start + 1)
                            fy = first_char.Information(6)
                            fx = first_char.Information(5)
                            cells.append({"r": ri, "c": ci, "y": round(fy, 2), "x": round(fx, 2)})
                        except Exception:
                            pass
                data["tables"].append({
                    "ti": ti,
                    "top_y": round(tbl_top, 2),
                    "page": tbl_page,
                    "cells": cells,
                })
            doc.Close(False)
            return data
        except Exception as e:
            last = e
            time.sleep(0.8 + attempt * 0.5)
    return {"doc": docx_path.name, "error": str(last)}


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        for doc_name in DOCS:
            path = DOCX_DIR / doc_name
            if not path.exists():
                results.append({"doc": doc_name, "error": "not found"})
                continue
            print(f"\n{doc_name}:")
            r = measure(word, path)
            results.append(r)
            if "error" in r:
                print(f"  ERR: {r['error']}")
                continue
            for tbl in r["tables"]:
                print(f"  Table {tbl['ti']} page {tbl['page']}: top_y={tbl['top_y']}")
                for c in tbl["cells"][:4]:
                    print(f"    R{c['r']}C{c['c']} y={c['y']} x={c['x']}")
    finally:
        try: word.Quit()
        except: pass
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
