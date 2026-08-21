# -*- coding: utf-8 -*-
"""How far below a table does Word start the next paragraph?

The last blocker for S1189. technical__00501ca3 page 8:
    Word  bottom rule y0 631.42 -> next paragraph glyph top 649.19 = 17.77
    Oxi   bottom rule y0 631.22 -> next paragraph glyph top 650.01 = 18.79
a +1.02pt over-advance that pushes '(Added 2002) (Amended 2010)' onto page 9
(1 paragraph of 272 — the doc's only pagination miss).

`_pb_tblborder_gen.py` already showed the foot gap tracks the bottom border's
DRAWN width (last drawn edge -> next glyph top: 0.79 / 1.27 / 1.87 / 3.19 / 6.19
for single sz 4/8/12/24/48, and 1.75 / 3.07 / 2.71 for double4 / double8 /
triple4), but there every edge of the table moved together. This probe varies
ONLY the bottom edge, and separately sweeps the following paragraph's
spacing-before, so the two terms can be read apart:

    foot advance = last row content -> next paragraph line
                 = (row box remainder) + bottom_drawn + para_before ?

  python tools/metrics/_pb_tblfoot_gen.py --measure --oxi
"""
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = Path(os.environ.get("OXI_SCRATCH", tempfile.gettempdir())) / "pb_tblfoot.docx"
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
N_ROWS = 3


def mark(i, side):
    return "ZMARK%s%02dZ" % (side, i)


def run(t):
    return ('<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
            '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>' % t)


def bd(tag, spec):
    if spec is None:
        return ""
    style, sz = spec
    if style == "nil":
        return '<w:%s w:val="nil"/>' % tag
    return '<w:%s w:val="%s" w:sz="%d" w:space="0" w:color="auto"/>' % (tag, style, sz)


def tbl(bottom, lastrow_cell_bottom=None, mixed=False):
    """`lastrow_cell_bottom` declares tcBorders bottom on the LAST row's cells.
    `mixed` declares it on only the FIRST of two cells (the real-doc shape:
    forms__0020466f has 10 of 18 tables whose last row nils its bottom while the
    table still declares `bottom single sz6`)."""
    edges = (bd("top", ("single", 4)) + bd("left", ("single", 4))
             + bd("right", ("single", 4)) + bd("insideH", ("single", 4))
             + bd("insideV", ("single", 4)) + bd("bottom", bottom))
    def cell(k, decl):
        tb = "<w:tcBorders>" + bd("bottom", decl) + "</w:tcBorders>" if decl else ""
        return ('<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>' + tb + "</w:tcPr>"
                '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                + run("ROW%d" % k) + "</w:p></w:tc>")
    rows = []
    for k in range(N_ROWS):
        d = lastrow_cell_bottom if k == N_ROWS - 1 else None
        d2 = None if (mixed and k == N_ROWS - 1) else d
        rows.append("<w:tr>" + cell(k, d) + cell(k, d2) + "</w:tr>")
    rows = "".join(rows)
    return ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
            '<w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar>'
            '<w:tblBorders>' + edges + "</w:tblBorders></w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' + rows + "</w:tbl>")


S4 = ("single", 4)
S8 = ("single", 8)
S12 = ("single", 12)
S24 = ("single", 24)
D4 = ("double", 4)
D8 = ("double", 8)
NIL = ("nil", 0)

# (name, table bottom border, following paragraph's spacing-before in twips)
ARMS = [
    ("bot_s4_b0", S4, None, None, False),
    ("bot_s8_b0", S8, None, None, False),
    ("bot_s12_b0", S12, None, None, False),
    ("bot_s24_b0", S24, None, None, False),
    ("bot_d4_b0", D4, None, None, False),
    ("bot_d8_b0", D8, None, None, False),
    ("bot_nil_b0", NIL, None, None, False),
    ("bot_s4_b200", S4, 200, None, False),
    ("bot_d4_b200", D4, 200, None, False),
    ("bot_s4_b400", S4, 400, None, False),
    # last row's cells override the table's bottom (the forms__0020466f shape)
    ("bot_s6_cellnil", ("single", 6), None, NIL, False),
    ("bot_s6_cellnil_mix", ("single", 6), None, NIL, True),
    ("bot_s6_cells24", ("single", 6), None, S24, False),
]
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def build():
    body = []
    for i, (name, bot, before, cellbot, mixed) in enumerate(ARMS):
        brk = '<w:p><w:pPr><w:pageBreakBefore/></w:pPr>' + run(mark(i, "A")) + "</w:p>" if i \
            else "<w:p>" + run(mark(i, "A")) + "</w:p>"
        body.append(brk)
        body.append(tbl(bot, cellbot, mixed))
        spc = '<w:spacing w:before="%d" w:after="0"/>' % before if before else '<w:spacing w:after="0"/>'
        body.append("<w:p><w:pPr>" + spc + "</w:pPr>" + run("AFTERTABLE") + "</w:p>")
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


def word_rows(path):
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
                    out.append((pi * 10000 + min(s["bbox"][1] for s in ln["spans"]), "T", t, 0.0))
        for dd in pg.get_drawings():
            rc = dd["rect"]
            if rc.height < 8 and rc.width > 20:
                out.append((pi * 10000 + rc.y0, "B", "", rc.height))
    return out


def oxi_rows(path):
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
                out.append((y, "T", (e.get("text") or "").strip(), 0.0))
            elif e.get("type") == "border" and (e.get("w") or 0) > 20:
                out.append((y, "B", "", 0.0))
    return out


def summarize(rows, tag):
    rows.sort()
    print("--- %s ---" % tag)
    res = {}
    for i, (name, bot, before, cellbot, mixed) in enumerate(ARMS):
        a = [y for y, k, t, h in rows if k == "T" and mark(i, "A") in t]
        b = [y for y, k, t, h in rows if k == "T" and mark(i, "B") in t]
        if not a or not b:
            print("  %-14s MISSING" % name)
            continue
        a, b = min(a), min(b)
        lastrow = [y for y, k, t, h in rows if k == "T" and t == "ROW%d" % (N_ROWS - 1)
                   and a < y < b]
        after = [y for y, k, t, h in rows if k == "T" and t == "AFTERTABLE" and a < y < b]
        edges = sorted({round(y, 2) for y, k, t, h in rows if k == "B" and a < y < b})
        merged = []
        for y in edges:
            if not merged or y - merged[-1] > 1.6:
                merged.append(y)
        if not lastrow or not after:
            print("  %-14s MISSING rows" % name)
            continue
        lr, af = min(lastrow), min(after)
        botedge = merged[-1] if merged else float("nan")
        print("  %-14s lastrow %8.2f  bottom_edge %8.2f  next %8.2f   row->next %6.2f   edge->next %6.2f"
              % (name, lr, botedge, af, af - lr, af - botedge))
        res[name] = (af - lr, af - botedge)
    return res


if __name__ == "__main__":
    p = build()
    w = summarize(word_rows(p), "WORD") if "--measure" in sys.argv else None
    o = summarize(oxi_rows(p), "OXI") if "--oxi" in sys.argv else None
    if w and o:
        print("--- DIFF (oxi - word) ---")
        for arm in ARMS:
            n = arm[0]
            if n in w and n in o:
                print("  %-14s d(row->next) %+7.2f   (word %.2f / oxi %.2f)"
                      % (n, o[n][0] - w[n][0], w[n][0], o[n][0]))
