# -*- coding: utf-8 -*-
"""How tall is the LINE that holds an inline image?

S536/S537 drop the image-only host paragraph and give the line the image's
extent; S965 carries the paragraph's spacing but not its line metrics. Two real
specimens bracket the answer without settling it:

  policies__0028d1be  a 7.5pt spacer.gif in a Times-New-Roman-12 cell
                      -> Word consumes 13.760pt border-to-image-bottom, so the
                         extent alone (7.5) is too small
  3a4f (calendar)     a 321.75pt drawing -> host + extent overshoots by 17.75pt,
                         so the sum is too big

`max(host_line, extent)` explains both, but so does a baseline composition
`max(ascent, extent) + descent`. They differ by the descent — about half a point
— so the discriminating case is an image just TALLER than the host line
(M = 18pt against a ~13.8pt line), where max says 18.0 and composition says
18 + descent.

Each variant is a 3-row table: a TOP marker row, the image-only row, a BOT
marker row, all in the same style so the readout is marker-baseline to
marker-baseline (never Word ink-top against an Oxi box-top). CTRL has no image
and measures the host line by itself.

  python _pb_imgline_gen.py gen       # build the variants from the real package
  python _pb_imgline_gen.py measure   # Word COM -> PDF
  python _pb_imgline_gen.py read      # per-variant line height + model verdict
"""
from __future__ import annotations

import glob
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_imgline"
HOST = REPO / "pipeline_data" / "docx_corpus" / "en" / "policies" / "0028d1bea47058b2.docx"

# extents in points; the image part is a 1x1 gif so the extent alone sets the size
VARIANTS = {"S": 7.5, "M": 18.0, "L": 321.75}

RPR = ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman"'
       ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
       '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>')
PPR = ('<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
       + RPR + '</w:pPr>')


def marker(text: str) -> str:
    return f'<w:p>{PPR}<w:r>{RPR}<w:t>{text}</w:t></w:r></w:p>'


def image_para(pt: float) -> str:
    cx, cy = 9525, int(round(pt * 12700))
    return (f'<w:p>{PPR}<w:r>{RPR}<w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="1" name="p"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="rIdIMG"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')


def row(inner: str) -> str:
    return ('<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/>'
            '<w:tcBorders>'
            '<w:top w:val="single" w:sz="8" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="8" w:color="000000"/>'
            '</w:tcBorders></w:tcPr>' + inner + '</w:tc></w:tr>')


def table(rows: str) -> str:
    return ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' + rows + '</w:tbl>')


def body_for(vid: str) -> str:
    """TOP marker row / image row / BOT marker row, plus a CTRL trio."""
    out = []
    # CTRL: same three rows with an EMPTY host paragraph instead of the image
    out.append(table(row(marker("CTRLTOP")) + row(marker("")) + row(marker("CTRLBOT"))))
    out.append(marker("---"))
    inner = image_para(VARIANTS[vid]) if vid in VARIANTS else marker("")
    out.append(table(row(marker(f"TOP{vid}")) + row(inner) + row(marker(f"BOT{vid}"))))
    return "".join(out)


def gen() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    src = zipfile.ZipFile(HOST)
    names = src.namelist()
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)
    sectpr = sect.group(0) if sect else ""
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    rels = src.read("word/_rels/document.xml.rels").decode("utf-8")
    # the image relationship id used by the real package
    m = re.search(r'Id="([^"]+)"[^>]*image1\.gif', rels)
    img_rid = m.group(1) if m else "rId4"
    for vid in VARIANTS:
        path = OUT / f"img_{vid}.docx"
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in names:
                if n == "word/document.xml":
                    body = body_for(vid).replace("rIdIMG", img_rid)
                    z.writestr(n, head + body + sectpr + "</w:body></w:document>")
                else:
                    z.writestr(n, src.read(n))
    print(f"gen: {len(VARIANTS)} variants (image rId={img_rid})")


def measure() -> None:
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        for path in sorted(glob.glob(str(OUT / "*.docx"))):
            pdf = path[:-5] + ".pdf"
            if os.path.exists(pdf):
                continue
            d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            try:
                d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf), ExportFormat=17)
            finally:
                d.Close(False)
            print("measured", os.path.basename(path), flush=True)
    finally:
        word.Quit()


def read() -> None:
    import fitz
    for vid in VARIANTS:
        pdf = OUT / f"img_{vid}.pdf"
        if not pdf.is_file():
            print(f"{vid}: no pdf"); continue
        doc = fitz.open(pdf)
        base = {}
        for page in doc:
            for b in page.get_text("dict")["blocks"]:
                for line in b.get("lines", []):
                    t = "".join(s["text"] for s in line["spans"]).strip()
                    if t.startswith(("CTRL", "TOP", "BOT")):
                        base.setdefault(t, round(line["spans"][0]["origin"][1], 3))
        ctrl = base.get("CTRLBOT", 0) - base.get("CTRLTOP", 0)
        var = base.get(f"BOT{vid}", 0) - base.get(f"TOP{vid}", 0)
        # both trios differ only in the middle row's content
        h = var - ctrl
        extent = VARIANTS[vid]
        print(f"  {vid}: CTRL span {ctrl:.3f}  variant span {var:.3f}  "
              f"image-line minus empty-line = {h:+.3f}  (extent {extent})")
        doc.close()
    print("\nempty host line L is the CTRL middle row; a `max(L, I)` model predicts")
    print("  delta = max(L, I) - L ;  a `max(ascent, I) + descent` model adds the descent.")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    {"gen": gen, "measure": measure, "read": read}[cmd]()
