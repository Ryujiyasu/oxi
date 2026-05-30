"""S448 step-2 starter: measure Word's ACTUAL table row pitch via COM and
compare to Oxi's, to derive the per-row shortfall rule (the step S183/208/223/445
skipped — they guessed a constant instead of measuring).

For each table in a doc, iterate tbl.Rows and read each row's TOP via the
R30-safe collapsed-start Information(6). Row pitch = top(next) - top(this).
Cross-reference trHeight/hRule (from XML) and report Word pitch per row.
Then load Oxi's pagination dump and report Oxi's row tops for the same table
(matched by the row's first cell text) so Word_pitch - Oxi_pitch is visible.

Usage (from repo root):
  python tools/metrics/_s448_rowpitch_com.py 34140b9c5662_index-14
"""
from __future__ import annotations
import io
import json
import os
import re
import sys
import zipfile

WD_VERT = 6  # wdVerticalPositionRelativeToPage
EMU = 1.0    # Information(6) returns points already


def xml_trheights(path):
    x = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "replace")
    dg = re.search(r"<w:docGrid[^>]*?/>", x)
    out = []
    for tr in re.findall(r"<w:trPr>(.*?)</w:trPr>", x, re.S):
        h = re.search(r'<w:trHeight[^>]*w:val="(\d+)"', tr)
        hr = re.search(r'w:hRule="(\w+)"', tr)
        out.append((int(h.group(1)) if h else None, hr.group(1) if hr else "auto"))
    return (dg.group(0) if dg else "none"), out


def measure_word(docx_abspath):
    import win32com.client as win32

    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    rows_out = []
    try:
        doc = app.Documents.Open(docx_abspath, ReadOnly=True)
        try:
            for ti in range(1, doc.Tables.Count + 1):
                tbl = doc.Tables(ti)
                prev_top = None
                for ri in range(1, tbl.Rows.Count + 1):
                    try:
                        rng = tbl.Rows(ri).Range
                        top = doc.Range(rng.Start, rng.Start).Information(WD_VERT)
                    except Exception:
                        top = None
                    # first-cell text
                    try:
                        c0 = tbl.Cell(ri, 1).Range.Text.strip("\r\x07 \t\n")
                    except Exception:
                        c0 = ""
                    pitch = (top - prev_top) if (top is not None and prev_top is not None) else None
                    rows_out.append({"tbl": ti, "row": ri, "top": top,
                                     "pitch_prev": pitch, "c0": c0[:14]})
                    if top is not None:
                        prev_top = top
            doc.Close(False)
        except Exception as e:
            doc.Close(False)
            raise
    finally:
        app.Quit()
    return rows_out


def oxi_rows(doc_id):
    p = f"pipeline_data/pagination_oxi/{doc_id}.json"
    if not os.path.exists(p):
        return []
    o = json.load(io.open(p, encoding="utf-8"))
    flat = []
    for pg, recs in o.get("pages", {}).items():
        for r in recs:
            if r.get("cell_col_idx") == 0 and r.get("y") is not None:
                flat.append((int(pg), r["y"], (r.get("text") or "")[:14]))
    return flat


def main():
    name = sys.argv[1]
    doc_id = name[:12]
    docx = os.path.abspath(f"tools/golden-test/documents/docx/{name}.docx")
    dg, trh = xml_trheights(docx)
    print(f"doc={name}  docGrid={dg}")
    print(f"explicit trHeight rows: {trh}")
    print("\n=== WORD actual row tops + pitch (COM) ===")
    wr = measure_word(docx)
    for r in wr:
        pp = f"{r['pitch_prev']:+6.2f}" if r["pitch_prev"] is not None else "   -- "
        tp = f"{r['top']:7.2f}" if r["top"] is not None else "  None "
        print(f"  tbl{r['tbl']} row{r['row']:2d} top={tp} pitch={pp}  c0={r['c0']!r}")


if __name__ == "__main__":
    main()
