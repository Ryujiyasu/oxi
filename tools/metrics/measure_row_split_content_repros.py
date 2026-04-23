"""Measure each RS_* repro via Word COM: for every table cell paragraph,
enumerate characters and record (offset, page, y_pt). Find the line
transitions (y jumps) and page transitions (page changes).

Output: pipeline_data/row_split_content_measurements.json
"""
import win32com.client
import json
import os
from pathlib import Path

REPRO_DIR = Path(__file__).parent / "row_split_content_repro"
OUT_JSON = Path("pipeline_data") / "row_split_content_measurements.json"


def measure_doc(word, docx_path: Path):
    """Open docx in Word, dump every char of every table cell paragraph."""
    doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
    result = {
        "file": docx_path.name,
        "total_chars": doc.Range().End,
        "page_count": doc.ComputeStatistics(2),  # wdStatisticPages
        "tables": [],
    }
    try:
        tables = doc.Tables
        for ti in range(1, tables.Count + 1):
            tbl = tables(ti)
            tinfo = {"index": ti, "rows": [], "cells_total": 0}
            for ri in range(1, tbl.Rows.Count + 1):
                row = tbl.Rows(ri)
                rinfo = {"row": ri, "cells": []}
                for ci in range(1, row.Cells.Count + 1):
                    cell = row.Cells(ci)
                    cinfo = {"col": ci, "paras": []}
                    pr = cell.Range
                    p_count = pr.Paragraphs.Count
                    for pi in range(1, p_count + 1):
                        para = pr.Paragraphs(pi)
                        s = para.Range.Start
                        e = para.Range.End
                        # sample chars within paragraph
                        lines = []
                        prev_y = None
                        prev_page = None
                        for off in range(s, e):
                            r = doc.Range(off, off + 1)
                            pg = r.Information(3)
                            y = r.Information(6)
                            if prev_y is None or abs(y - prev_y) > 0.3 or pg != prev_page:
                                ch = r.Text[:1].replace('\r', '¶').replace('\n', '↵').replace('\t', '→')
                                if ch == '\x07':
                                    ch = '⌂'  # cell-end marker
                                lines.append({
                                    "offset": off,
                                    "page": int(pg),
                                    "y_pt": round(y, 2),
                                    "char": ch,
                                })
                                prev_y = y
                                prev_page = pg
                        cinfo["paras"].append({
                            "p_idx": pi,
                            "start": s,
                            "end": e,
                            "line_count": len(lines),
                            "lines": lines,
                        })
                    rinfo["cells"].append(cinfo)
                    tinfo["cells_total"] += 1
                tinfo["rows"].append(rinfo)
            result["tables"].append(tinfo)
        doc.Close(SaveChanges=False)
    except Exception as ex:
        try:
            doc.Close(SaveChanges=False)
        except Exception:
            pass
        result["error"] = str(ex)
    return result


def main():
    files = sorted(REPRO_DIR.glob("RS_*.docx"))
    if not files:
        print(f"No RS_*.docx files in {REPRO_DIR}")
        return
    print(f"Found {len(files)} repros")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    out = []
    try:
        for f in files:
            print(f"\n=== {f.name} ===")
            r = measure_doc(word, f)
            print(f"  pages={r['page_count']} tables={len(r['tables'])}")
            for t in r["tables"]:
                for row in t["rows"]:
                    for c in row["cells"]:
                        for p in c["paras"]:
                            if p["line_count"] <= 1:
                                continue
                            pages_seen = sorted({ln["page"] for ln in p["lines"]})
                            print(f"    tbl{t['index']} row{row['row']} col{c['col']} para{p['p_idx']}  "
                                  f"lines={p['line_count']}  pages={pages_seen}")
                            # If multi-page, show the split point
                            if len(pages_seen) >= 2:
                                for i, ln in enumerate(p["lines"]):
                                    mark = ""
                                    if i > 0 and ln["page"] != p["lines"][i-1]["page"]:
                                        mark = "  ** SPLIT **"
                                    print(f"      line{i} off={ln['offset']:5d} p{ln['page']} "
                                          f"y={ln['y_pt']:7.2f} ch={ln['char']!r}{mark}")
            out.append(r)
    finally:
        word.Quit()

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nSaved → {OUT_JSON}")


if __name__ == "__main__":
    main()
