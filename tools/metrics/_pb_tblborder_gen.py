# -*- coding: utf-8 -*-
"""Derive how much a table BORDER adds to a body row's box and to the flow
BELOW the table, as a function of (style, sz).

Lead: forms__000ee7c0 (EN pcd census, Word 2 pages / Oxi 1). Its page-1 body
ends 8.2pt SHORT of Word's, which is why Oxi keeps the trailing empty paragraph
that Word pushes to page 2. Walking the border ladder localizes ALL of it to
border draw width:

    Word row tops  21.60 31.20 41.64 52.44 63.24 74.04   pitches 9.60 10.44 10.80 10.80 10.80
    Oxi  row tops  21.50 30.05 38.86 49.66 60.46 71.26   pitches 8.55  8.81 10.80 10.80 10.80

The two rows that differ are exactly the two bounded by a DOUBLE border (Word
draws companion lines at 22.56/23.04 and 32.16/32.64/32.88); every row bounded
by SINGLE borders matches to 0.00. Same at the table's foot: Word's outer double
bottom occupies 74.04..75.72 and the next paragraph starts below it, Oxi charges
zero. The footer stack already models this (`eff_bw`: double = 3x sz/8 per side,
mod.rs S868b) — the BODY row path has no `double` arm at all.

Each arm is one table (4 identical one-line rows) bracketed by hyphen-free
markers, so the PDF gives both the row pitches and the post-table advance.

  python tools/metrics/_pb_tblborder_gen.py            # write the probe docx
  python tools/metrics/_pb_tblborder_gen.py --measure  # + Word PDF truth
  python tools/metrics/_pb_tblborder_gen.py --oxi      # + Oxi's answer
"""
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = Path(os.environ.get("OXI_SCRATCH", tempfile.gettempdir())) / "pb_tblborder.docx"

NS = ' '.join([
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
])

# (name, border style, sz eighths-of-a-point)
ARMS = [
    ("none", None, 0),
    ("single4", "single", 4),
    ("single8", "single", 8),
    ("single12", "single", 12),
    ("single24", "single", 24),
    ("single48", "single", 48),
    ("double4", "double", 4),
    ("double8", "double", 8),
    ("double12", "double", 12),
    ("dashed4", "dashed", 4),
    ("thick4", "thick", 4),
    ("triple4", "triple", 4),
]
N_ROWS = 4


def mark(i, side):
    return "ZMARK%s%02dZ" % (side, i)


def txt(s, sz=20):
    return ('<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>' % (sz, sz, s))


def tbl(style, sz):
    if style is None:
        borders = ""
    else:
        edge = ('<w:%%s w:val="%s" w:sz="%d" w:space="0" w:color="auto"/>' % (style, sz))
        borders = ("<w:tblBorders>"
                   + "".join(edge % k for k in ("top", "left", "bottom", "right",
                                                "insideH", "insideV"))
                   + "</w:tblBorders>")
    rows = "".join(
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        + txt("R%d" % k) + "</w:p></w:tc>"
        '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        + txt("C%d" % k) + "</w:p></w:tc></w:tr>"
        for k in range(N_ROWS))
    return ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/>' + borders + "</w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            + rows + "</w:tbl>")


SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def build():
    body = []
    for i, (name, style, sz) in enumerate(ARMS):
        # ★One arm per PAGE. Without the break an arm that straddles a page
        # boundary reports a ~9400pt advance and a nonsense pitch list, and a
        # 6pt border is thick enough to shift the arms after it.
        brk = ('<w:pPr><w:pageBreakBefore/></w:pPr>' if i else "")
        body.append("<w:p>" + brk + txt(mark(i, "A")) + "</w:p>")
        body.append(tbl(style, sz))
        body.append("<w:p>" + txt(mark(i, "B")) + "</w:p>")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (NS, "".join(body), SECT))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          "</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            "</Relationships>")
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
            # ★height < 8, not < 3: a sz24/sz48 edge is 3pt/6pt thick and a
            # `< 3` filter dropped it entirely, so those arms reported ZERO
            # edges and looked unmeasurable.
            if rc.height < 8 and rc.width > 20:
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
            elif e.get("type") == "border" and (e.get("w") or 0) > 20:
                out.append((y, "B", ""))
    return out


def summarize(rows, tag):
    rows.sort()
    print("--- %s ---" % tag)
    res = {}
    for i, (name, _s, _z) in enumerate(ARMS):
        ya = [r[0] for r in rows if r[1] == "T" and mark(i, "A") in r[2]]
        yb = [r[0] for r in rows if r[1] == "T" and mark(i, "B") in r[2]]
        if not ya or not yb:
            print("  %-9s MISSING" % name)
            continue
        a, b = min(ya), min(yb)
        bs = sorted({round(r[0], 2) for r in rows if r[1] == "B" and a < r[0] < b})
        # merge companion lines of one logical edge (< 2pt apart)
        edges = []
        for y in bs:
            if not edges or y - edges[-1] > 2.0:
                edges.append(y)
        pitches = [round(edges[k + 1] - edges[k], 2) for k in range(len(edges) - 1)]
        print("  %-9s advance %7.2f  edges %d  pitches %s  last_edge->B %6.2f"
              % (name, b - a, len(edges), pitches, (b - edges[-1]) if edges else float("nan")))
        res[name] = (b - a, pitches, (b - edges[-1]) if edges else None)
    return res


if __name__ == "__main__":
    p = build()
    w = summarize(rows_word(p), "WORD") if "--measure" in sys.argv else None
    o = summarize(rows_oxi(p), "OXI") if "--oxi" in sys.argv else None
    if w and o:
        print("--- DIFF (oxi - word) ---")
        for name, _s, _z in ARMS:
            if name in w and name in o:
                print("  %-9s d_advance %+7.2f   pitch_w %s  pitch_o %s"
                      % (name, o[name][0] - w[name][0], w[name][1], o[name][1]))
