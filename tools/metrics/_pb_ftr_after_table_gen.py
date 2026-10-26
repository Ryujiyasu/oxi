# -*- coding: utf-8 -*-
"""Does a footer paragraph that follows a table keep its space before?

en/technical__0061c884: footer3 = an empty Footer paragraph (Normal: before 2,
after 2), a one-row table (42pt picture, top border 1pt), an empty Footer
paragraph. Word stacks it at 80.30 above the footer distance (table top border
742.90); Oxi's footer stack is 78.30 (744.85) -- the paragraph after the table
lost its 2pt before. Faithful slices, the trailing paragraph's before changed:
    asis    inherited before 2
    b0      w:before="0"
    b12     w:before="240"
    b24     w:before="480"
    a12     the LEADING paragraph's after 12 (w:after="240"), trailing as is
Read: the top of the footer's 42pt picture on page 7 (3pt below the table's
top border in both Word and Oxi).

    python tools/metrics/_pb_ftr_after_table_gen.py   (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import json, os, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/technical/0061c884d3e2d802.docx"
OUT = REPO / "pipeline_data" / "_pb_ftr_after_table"
PAGE = 6  # page 7: the first page whose footer is footer3 (the table)
ARMS = ["asis", "b0", "b12", "b24", "a12"]
EMPTY = '<w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'


def variant(name):
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    f = parts["word/footer3.xml"].decode("utf-8")
    i = f.find("</w:tbl>")
    head, tail = f[:i], f[i:]
    assert tail.count(EMPTY) == 1 and head.count(EMPTY) == 1
    if name in ("b0", "b12", "b24"):
        tw = {"b0": 0, "b12": 240, "b24": 480}[name]
        tail = tail.replace(EMPTY, f'<w:pPr><w:pStyle w:val="Footer"/><w:spacing w:before="{tw}"/></w:pPr></w:p>')
    elif name == "a12":
        head = head.replace(EMPTY, '<w:pPr><w:pStyle w:val="Footer"/><w:spacing w:after="240"/></w:pPr></w:p>')
    parts["word/footer3.xml"] = (head + tail).encode("utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


def word_top(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[PAGE]
    ys = [b["bbox"][1] for b in pg.get_text("dict")["blocks"]
          if b.get("type") == 1 and abs(b["bbox"][3] - b["bbox"][1] - 42) < 0.5]
    return max(ys) if ys else None


def oxi_top(exe, docx):
    t = tempfile.mkdtemp(); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "q"), "--dump-layout=" + lay], capture_output=True)
    pg = json.load(open(lay, encoding="utf-8"))["pages"][PAGE]
    ys = [e["y"] for e in pg["elements"] if e.get("type") == "image" and abs(e.get("h", 0) - 42) < 0.5]
    return max(ys) if ys else None


if __name__ == "__main__":
    exes = [e for e in os.environ.get("PB_OXI", "").split(";") if e]
    for name in os.environ.get("PB_ARMS", ",".join(ARMS)).split(","):
        p = variant(name); pdf = p.with_suffix(".pdf")
        if not pdf.exists():
            import win32com.client
            w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
            try:
                d = w.Documents.Open(str(p), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
            finally:
                w.Quit()
        wt = word_top(pdf)
        cells = [f"{name:5} W {wt:7.2f}"]
        for e in exes:
            ot = oxi_top(e, p)
            cells.append(f"O {ot:7.2f} ({ot - wt:+.2f})" if ot is not None else "O -")
        print("   ".join(cells), flush=True)
