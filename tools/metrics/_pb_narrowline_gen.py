# -*- coding: utf-8 -*-
"""Is Oxi's line height for Arial Narrow the FONT's, or a 1.2x fallback?

Decomposition of technical__0009d767 (the doc S865's carve-out was built for):

    Word interior row 20.88 = cellMar 3.6+3.6 + line 13.20
    Oxi  interior row 21.00 = cellMar 3.6+3.6 + line 13.80

    line / (276/240)  ->  Word 11.478   Oxi 12.000
    Arial Narrow hhea (asc 1916, desc -434, gap 0, upm 2048) @10pt = 11.4746
    generic 1.2 * size                                             = 12.0000

So Word uses the face's own hhea and Oxi looks to be on the 1.2x fallback —
even though `Arial Narrow` IS present in font_metrics_compact.json. This probe
separates the two candidate causes: a metrics LOOKUP that misses, or a CELL
path that never consults hhea. Each arm appears twice, once in body text and
once in a one-cell table, so the two answers can differ.

Arial control arms are included because Arial and Arial Narrow have DIFFERENT
hhea sums (11.4990 vs 11.4746) — close, but the fallback 12.0 is far from both.

  python tools/metrics/_pb_narrowline_gen.py --measure --oxi
"""
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = Path(os.environ.get("OXI_SCRATCH", tempfile.gettempdir())) / "pb_narrowline.docx"
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

# (name, family, half-points, w:line or None, in_cell)
ARMS = [
    ("body_narrow_276", "Arial Narrow", 20, 276, False),
    ("body_arial_276", "Arial", 20, 276, False),
    ("body_narrow_auto", "Arial Narrow", 20, None, False),
    ("body_arial_auto", "Arial", 20, None, False),
    ("cell_narrow_276", "Arial Narrow", 20, 276, True),
    ("cell_arial_276", "Arial", 20, 276, True),
    ("cell_narrow_auto", "Arial Narrow", 20, None, True),
    ("cell_arial_auto", "Arial", 20, None, True),
    ("body_narrow_276_sz24", "Arial Narrow", 24, 276, False),
    ("cell_narrow_276_sz24", "Arial Narrow", 24, 276, True),
]
N_LINES = 4


def mark(i, side):
    return "ZMARK%s%02dZ" % (side, i)


def run(text, fam, hp):
    return ('<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>' % (fam, fam, fam, hp, hp, text))


def para(text, fam, hp, line):
    sp = ('<w:spacing w:after="0" w:line="%d" w:lineRule="auto"/>' % line if line
          else '<w:spacing w:after="0"/>')
    return "<w:p><w:pPr>" + sp + "</w:pPr>" + run(text, fam, hp) + "</w:p>"


def block(i, fam, hp, line, in_cell):
    """N_LINES separate paragraphs so the PITCH between them is the line box."""
    paras = "".join(para("L%d" % k, fam, hp, line) for k in range(N_LINES))
    if not in_cell:
        return paras
    # one cell, zero cell margins so the row pitch IS the stacked line boxes
    return ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
            '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
            "</w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>'
            + paras + "</w:tc></w:tr></w:tbl>")


SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def build():
    body = []
    for i, (name, fam, hp, line, in_cell) in enumerate(ARMS):
        brk = '<w:pPr><w:pageBreakBefore/></w:pPr>' if i else ""
        body.append("<w:p>" + brk + run(mark(i, "A"), "Calibri", 20) + "</w:p>")
        body.append(block(i, fam, hp, line, in_cell))
        body.append("<w:p>" + run(mark(i, "B"), "Calibri", 20) + "</w:p>")
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


def word_lines(path):
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
    out, fonts = [], set()
    for pi in range(doc.page_count):
        for blk in doc[pi].get_text("dict")["blocks"]:
            if blk.get("type", 0) != 0:
                continue
            for ln in blk.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if not t:
                    continue
                for s in ln["spans"]:
                    if s["text"].strip():
                        fonts.add(s["font"])
                out.append((pi * 10000 + min(s["bbox"][1] for s in ln["spans"]), t))
    # ★the truth is only the truth if Word really used the face we asked for
    print("   PDF fonts:", sorted(fonts))
    return out


def oxi_lines(path):
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
            if e.get("type") == "text" and (e.get("text") or "").strip():
                out.append((pi * 10000 + e.get("y", 0.0), (e.get("text") or "").strip()))
    return out


def summarize(rows, tag):
    rows.sort()
    print("--- %s ---" % tag)
    res = {}
    for i, (name, fam, hp, line, in_cell) in enumerate(ARMS):
        ys = sorted({round(y, 2) for y, t in rows
                     if t in ("L0", "L1", "L2", "L3")
                     and min(a for a, s in rows if mark(i, "A") in s) < y
                     < min(b for b, s in rows if mark(i, "B") in s)})
        if len(ys) < 2:
            print("  %-22s MISSING" % name)
            continue
        pitches = [round(ys[k + 1] - ys[k], 3) for k in range(len(ys) - 1)]
        avg = sum(pitches) / len(pitches)
        mult = (line / 240.0) if line else 1.0
        print("  %-22s pitches %-26s mean %.3f   natural=%.4f"
              % (name, str(pitches), avg, avg / mult))
        res[name] = avg
    return res


if __name__ == "__main__":
    p = build()
    w = summarize(word_lines(p), "WORD") if "--measure" in sys.argv else None
    o = summarize(oxi_lines(p), "OXI") if "--oxi" in sys.argv else None
    if w and o:
        print("--- DIFF (oxi - word) ---")
        for arm in ARMS:
            n = arm[0]
            if n in w and n in o:
                print("  %-22s %+7.3f   (word %.3f / oxi %.3f)" % (n, o[n] - w[n], w[n], o[n]))
