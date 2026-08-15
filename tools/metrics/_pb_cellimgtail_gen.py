# -*- coding: utf-8 -*-
"""How far does Word fill a page with a split cell's tail after a tall image?

educational__002a301d (0.6846, pcd +2): a table cell holds a tall image followed
by a run of empty paragraphs. Word keeps filling page N with those trailing
lines down to its content bottom (p5: image ends ~573, empties at 573.75 …
702.0) and starts page N+1 with the next real paragraph. Oxi stops at the
image's bottom (569.5) and re-emits SIX 16pt lines at the top of the
continuation, which pushes a later row off the page and costs the doc 2 pages.

The tail lines are invisible in a PDF when they are empty, so each carries a
short marker instead — same paragraph style, same line height, but readable.
Each arm shifts the table down by a different number of filler lines, so the
image bottom lands at a different distance from the page bottom, and the arm
reports which markers Word left on the first page.

  python _pb_cellimgtail_gen.py gen
  python _pb_cellimgtail_gen.py pdf      # Word truth
  python _pb_cellimgtail_gen.py oxi      # Oxi, same arms
"""
import json
import os
import struct
import subprocess
import sys
import tempfile
import zipfile
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cellimgtail")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

PGW, PGH, MARG = 12240, 15840, 1440      # Letter, 1in margins
IMG_H_PT = 220.0                          # the specimen's images are 226-271pt
IMG_W_PT = 340.0
TAIL = 9                                  # marker lines after the image
# ★The first cut swept 0-14 filler lines and NOTHING crossed a page: the line
# pitch is 13.8, so even arm 7 ended at 609 with the whole row on one page.
# The row (image 220 + 9 tail lines 124 = 344) only crosses when its top is
# past 720-344 = 376, i.e. from ~22 filler lines on; the sweep now walks the
# boundary through the tail.
FILLERS = [22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44]


def docx():
    return os.path.join(OUT, "cellimgtail.docx")


def png_bytes(w=340, h=220):
    """A minimal opaque PNG, generated so no binary asset is checked in."""
    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c))
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    row = b"\x00" + b"\x80\x80\x80" * w
    idat = zlib.compress(row * h, 6)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", idat) + chunk(b"IEND", b""))


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="24"/></w:rPr>')


def para(text, pbb=False):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/>%s</w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t>'
            "</w:r></w:p>" % ("<w:pageBreakBefore/>" if pbb else "", rpr(), rpr(), text))


def img_para():
    cx, cy = int(IMG_W_PT * 12700), int(IMG_H_PT * 12700)
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:drawing>'
            '<wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:docPr id="1" name="p"/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
            % (cx, cy, cx, cy))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, fill in enumerate(FILLERS):
        body.append(para("M%02d" % ai, pbb=ai > 0))
        for k in range(fill):
            body.append(para("a%df%d" % (ai, k)))
        # A 2-cell row (multi-cell = the specimen's shape, and the split path
        # S754 takes for it), the left cell carrying image + marker tail.
        cell = img_para() + "".join(para("a%dL%d" % (ai, k + 1)) for k in range(TAIL))
        body.append(
            '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/>'
            '<w:left w:val="single" w:sz="4" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="4" w:color="000000"/>'
            '<w:right w:val="single" w:sz="4" w:color="000000"/>'
            '<w:insideH w:val="single" w:sz="4" w:color="000000"/>'
            '<w:insideV w:val="single" w:sz="4" w:color="000000"/></w:tblBorders>'
            "</w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="7000"/><w:gridCol w:w="2360"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>'
            + cell + "</w:tc>"
            '<w:tc><w:tcPr><w:tcW w:w="2360" w:type="dxa"/></w:tcPr>'
            + para("side") + "</w:tc></w:tr></w:tbl>")
        body.append(para("E%02d" % ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="%d" w:h="%d"/>'
           '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (PGW, PGH, MARG, MARG, MARG, MARG))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
              '<w:sz w:val="24"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    ct = CT.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>')
    drels = DRELS.replace("</Relationships>",
                          '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/image" Target="media/p.png"/>'
                          "</Relationships>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/media/p.png", png_bytes())
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(FILLERS), "arms | image", IMG_H_PT, "pt | tail", TAIL)


def report(per, who):
    print("== %s ==  (content bottom = %.1fpt)" % (who, (PGH - MARG) / 20.0))
    print("%-6s %-8s %-28s %s" % ("arm", "fillers", "tail on page 1", "img bottom"))
    for ai, fill in enumerate(FILLERS):
        got = per.get(ai)
        if not got:
            print("%-6d %-8d MISSING" % (ai, fill))
            continue
        kept, imgb = got
        print("%-6d %-8d %-28s %s"
              % (ai, fill, " ".join(kept) if kept else "(none)",
                 "%.1f" % imgb if imgb else "?"))


def pdf():
    import fitz
    import win32com.client as w
    out = docx().replace(".docx", ".pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    per = {}
    page_of = {}
    for pi in range(doc.page_count):
        for bl in doc[pi].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t:
                    page_of.setdefault(t, (pi, round(ln["bbox"][1], 1)))
    for ai in range(len(FILLERS)):
        start = page_of.get("M%02d" % ai)
        if not start:
            continue
        p0 = start[0]
        kept = [f"L{k}" for k in range(1, TAIL + 1)
                if page_of.get(f"a{ai}L{k}", (-1,))[0] == p0]
        # image bottom on that page = the tallest drawing rect
        imgb = None
        for dr in doc[p0].get_drawings():
            r = dr["rect"]
            # the IMAGE, not the row's tall left border: match its width too
            if r.height > 100 and abs(r.width - IMG_W_PT) < 12 and (imgb is None or r.y1 > imgb):
                imgb = r.y1
        per[ai] = (kept, imgb)
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "cellimgtail_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "cit"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    page_of = {}
    imgs = {}
    for pi, pg in enumerate(pages):
        for e in pg["elements"]:
            t = (e.get("text") or "").strip()
            if t:
                page_of.setdefault(t, (pi, round(e["y"], 1)))
            if e.get("type") == "image" and (e.get("h") or 0) > 100:
                imgs.setdefault(pi, e["y"] + e["h"])
    per = {}
    for ai in range(len(FILLERS)):
        start = page_of.get("M%02d" % ai)
        if not start:
            continue
        p0 = start[0]
        kept = [f"L{k}" for k in range(1, TAIL + 1)
                if page_of.get(f"a{ai}L{k}", (-1,))[0] == p0]
        per[ai] = (kept, imgs.get(p0))
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
