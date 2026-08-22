# -*- coding: utf-8 -*-
"""Where does a vMerge span put its EXCESS height — the first row or the last?

Blocker for S1189. technical__00501ca3 page-8 table (7 rows, two 2-row vMerge
spans), per-row pitch Word vs Oxi:

    r0 28.80/28.80  r1 40.32/40.30  r2 16.32/30.10  r3 22.95/16.55
    r4 16.32/16.30  r5 23.04/16.55

r2/r3 and r4/r5 are the spans. Word keeps the span's FIRST row at its own
one-line height (16.32) and lets the LAST row absorb the merged cell's
remainder (22.95 = 16.32 + 6.63); Oxi does the opposite, piling the remainder
onto the first row (30.10) and leaving the last at one line.

Each arm is one table on its own page: a 2- or 3-row vMerge span whose merged
cell carries N lines while every other cell carries exactly one, so the excess
is known and its placement is directly visible in the pitch list.

  python tools/metrics/_pb_vmergedist_gen.py --measure --oxi
"""
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
OUT = Path(os.environ.get("OXI_SCRATCH", tempfile.gettempdir())) / "pb_vmergedist.docx"
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def mark(i, side):
    return "ZMARK%s%02dZ" % (side, i)


def run(t):
    return ('<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
            '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>' % t)


def para(t):
    return ('<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            + run(t) + "</w:p>")


def cell(w, paras, vmerge=None):
    vm = ""
    if vmerge == "restart":
        vm = '<w:vMerge w:val="restart"/>'
    elif vmerge == "cont":
        vm = "<w:vMerge/>"
    return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>' % w + vm + "</w:tcPr>"
            + ("".join(para(p) for p in paras) if paras else para("")) + "</w:tc>")


def tbl(span_rows, merged_lines):
    """A `span_rows`-row vMerge span; the merged cell holds `merged_lines`
    single-word paragraphs, every other cell exactly one."""
    edges = "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>' % k
                    for k in ("top", "left", "bottom", "right", "insideH", "insideV"))
    rows = []
    for k in range(span_rows):
        if k == 0:
            merged = cell(4000, ["M%d" % n for n in range(merged_lines)], "restart")
        else:
            merged = cell(4000, None, "cont")
        rows.append("<w:tr>" + merged + cell(4000, ["R%d" % k]) + "</w:tr>")
    return ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
            '<w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar>'
            '<w:tblBorders>' + edges + "</w:tblBorders></w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            + "".join(rows) + "</w:tbl>")


# (name, rows in the span, lines in the merged cell)
ARMS = [
    ("span2_m1", 2, 1),
    ("span2_m2", 2, 2),
    ("span2_m3", 2, 3),
    ("span2_m4", 2, 4),
    ("span3_m1", 3, 1),
    ("span3_m3", 3, 3),
    ("span3_m5", 3, 5),
    ("span3_m7", 3, 7),
]
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def build():
    body = []
    for i, (name, sr, ml) in enumerate(ARMS):
        brk = '<w:p><w:pPr><w:pageBreakBefore/></w:pPr>' + run(mark(i, "A")) + "</w:p>" if i \
            else "<w:p>" + run(mark(i, "A")) + "</w:p>"
        body.append(brk)
        body.append(tbl(sr, ml))
        body.append("<w:p>" + run(mark(i, "B")) + "</w:p>")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, "".join(body), SECT))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
    print("wrote", OUT)
    return OUT


def rows_word(path):
    pdf = Path(tempfile.gettempdir()) / (path.stem + ".truth.pdf")
    if not pdf.exists() or "--reexport" in sys.argv:
        import win32com.client as win32
        w = win32.DispatchEx("Word.Application")
        w.Visible = False
        try:
            d = w.Documents.Open(str(path), ReadOnly=True)
            d.ExportAsFixedFormat(str(pdf), 17)
            d.Close(False)
        finally:
            w.Quit()
    import fitz
    doc = fitz.open(pdf)
    out = []
    for pi in range(doc.page_count):
        pg = doc[pi]
        for blk in pg.get_text("dict")["blocks"]:
            if blk.get("type", 0) != 0:
                continue
            for ln in blk.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t:
                    out.append((pi * 10000 + min(s["bbox"][1] for s in ln["spans"]), "T", t))
        for dd in pg.get_drawings():
            rc = dd["rect"]
            if rc.height < 8 and rc.width > 40:
                out.append((pi * 10000 + rc.y0, "B", ""))
    return out


def rows_oxi(path):
    exe = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
    tmp = Path(tempfile.mkdtemp())
    dump = tmp / "d.json"
    subprocess.run([str(exe), str(path), str(tmp / "p"), "110", "--dump-layout=%s" % dump],
                   check=True, capture_output=True)
    import json
    d = json.load(open(dump, encoding="utf-8"))
    out = []
    for pi, pg in enumerate(d["pages"]):
        for e in pg["elements"]:
            y = pi * 10000 + e.get("y", 0.0)
            if e.get("type") == "text" and (e.get("text") or "").strip():
                out.append((y, "T", (e.get("text") or "").strip()))
            elif e.get("type") == "border" and (e.get("w") or 0) > 40:
                out.append((y, "B", ""))
    return out


def summarize(rows, tag):
    rows.sort()
    print("--- %s ---" % tag)
    res = {}
    for i, (name, sr, ml) in enumerate(ARMS):
        a = [y for y, k, t in rows if k == "T" and mark(i, "A") in t]
        b = [y for y, k, t in rows if k == "T" and mark(i, "B") in t]
        if not a or not b:
            print("  %-10s MISSING" % name)
            continue
        a, b = min(a), min(b)
        ys = sorted({round(y, 2) for y, k, t in rows if k == "B" and a < y <= b + 0.5})
        m = []
        for y in ys:
            if not m or y - m[-1] > 1.6:
                m.append(y)
        p = [round(m[k + 1] - m[k], 2) for k in range(len(m) - 1)]
        print("  %-10s span=%d merged_lines=%-2d pitches %s" % (name, sr, ml, p))
        res[name] = p
    return res


if __name__ == "__main__":
    p = build()
    w = summarize(rows_word(p), "WORD") if "--measure" in sys.argv else None
    o = summarize(rows_oxi(p), "OXI") if "--oxi" in sys.argv else None
    if w and o:
        print("--- DIFF ---")
        for arm in ARMS:
            n = arm[0]
            if n in w and n in o and len(w[n]) == len(o[n]):
                print("  %-10s d=%s" % (n, [round(b - a, 2) for a, b in zip(w[n], o[n])]))
            elif n in w and n in o:
                print("  %-10s ROW COUNT word %d / oxi %d" % (n, len(w[n]), len(o[n])))
