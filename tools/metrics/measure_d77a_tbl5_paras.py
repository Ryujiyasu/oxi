"""Measure every paragraph inside d77a tbl5 (the table that splits p.6/p.7).
Dump per-paragraph line count + each line's (offset, page, y_pt) so we can
compare against Oxi's layout.

Also dump the page_bottom on p.6 and the cursor_y where the table starts.
"""
import win32com.client
import json
import os
from pathlib import Path

DOCX = Path("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx").resolve()
OUT = Path("pipeline_data/d77a_tbl5_word_measurements.json")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    out = {"file": DOCX.name, "tables": []}
    try:
        doc = word.Documents.Open(str(DOCX), ReadOnly=True)
        tables = doc.Tables
        # The table that splits is tbl5 (1-indexed in Word)
        TARGET_TBL = 5
        tbl = tables(TARGET_TBL)
        print(f"Table {TARGET_TBL}: rows={tbl.Rows.Count} cols={tbl.Columns.Count}")
        t_info = {"index": TARGET_TBL, "rows": tbl.Rows.Count, "cols": tbl.Columns.Count, "cells": []}

        for ri in range(1, tbl.Rows.Count + 1):
            row = tbl.Rows(ri)
            for ci in range(1, row.Cells.Count + 1):
                cell = row.Cells(ci)
                cell_range = cell.Range
                print(f"\n  Cell row{ri} col{ci}: chars {cell_range.Start}..{cell_range.End}, paras={cell_range.Paragraphs.Count}")
                c_info = {"row": ri, "col": ci, "start": cell_range.Start, "end": cell_range.End, "paras": []}

                for pi in range(1, cell_range.Paragraphs.Count + 1):
                    para = cell_range.Paragraphs(pi)
                    pr = para.Range
                    # Get the first char's text
                    preview = pr.Text[:40].replace('\r', '¶').replace('\x07', '⌂')
                    # Enumerate each char for line boundaries
                    lines = []
                    prev_y = None
                    prev_page = None
                    for off in range(pr.Start, pr.End):
                        r = doc.Range(off, off + 1)
                        pg = r.Information(3)
                        y = r.Information(6)
                        if prev_y is None or abs(y - prev_y) > 0.3 or pg != prev_page:
                            ch = r.Text[:1].replace('\r', '¶').replace('\n', '↵').replace('\t', '→')
                            if ch == '\x07':
                                ch = '⌂'
                            lines.append({"offset": off, "page": int(pg), "y_pt": round(y, 2), "char": ch})
                            prev_y = y
                            prev_page = pg
                    p_info = {
                        "p_idx": pi,
                        "start": pr.Start,
                        "end": pr.End,
                        "chars": pr.End - pr.Start,
                        "preview": preview,
                        "line_count": len(lines),
                        "lines": lines,
                    }
                    c_info["paras"].append(p_info)
                    pages_seen = sorted({ln["page"] for ln in lines})
                    print(f"    para{pi} chars={p_info['chars']:3d} lines={len(lines)} pages={pages_seen}  {preview[:25]!r}")
                    for ln in lines:
                        print(f"      off={ln['offset']:5d} p{ln['page']} y={ln['y_pt']:7.2f} ch={ln['char']!r}")

                t_info["cells"].append(c_info)
        out["tables"].append(t_info)

        # Also measure: what's the Y of the first body paragraph AFTER tbl5?
        # (= where the next element lands, useful to figure out end of table on p.6)
        print("\n--- paras just after tbl5 on p.6/p.7 boundary ---")
        after_off = t_info["cells"][-1]["end"]
        for off in range(after_off, after_off + 10):
            r = doc.Range(off, off + 1)
            pg = r.Information(3)
            y = r.Information(6)
            ch = r.Text[:1].replace('\r', '¶')
            if ch == '\x07':
                ch = '⌂'
            print(f"  off={off} p{pg} y={y:.2f} ch={ch!r}")
            if off > after_off + 4:
                break

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nSaved → {OUT}")


if __name__ == "__main__":
    main()
