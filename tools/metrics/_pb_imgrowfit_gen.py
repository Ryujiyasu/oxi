# -*- coding: utf-8 -*-
"""When a row's image CANNOT fit, may anything above it stay behind?

`_pb_cellimgtail_gen.py` settled the other half: while the image FITS, Word
keeps it and packs the tail (3/2/1/0 lines as it walks down). S1168 implements
that and matches 7/7. What it does not cover is the case that still costs
educational__00161422 its PASS:

  p12/p13: Word leaves 230pt of p12 BLANK and moves the row whole, although the
  caption above the image would fit; Oxi keeps 2 lines. No image crosses the
  break line there (they sit wholly below it), so the S1168 pull-back never
  fires.

and the three real documents disagree about it:

  technical__0061c884   image moves, NINE paragraphs above it stay behind
  educational__002a301d image moves, the six leading empties stay behind
  educational__00161422 image moves, NOTHING stays -- the whole row goes

So: put k lines ABOVE the image, make the image the LAST block (row overflow
then means exactly "the image does not fit", with no tail to confound it), and
walk the row down a line at a time. Read which of the k markers Word leaves on
the first page.

  k lines kept => keep-all-that-fit, and 00161422's push has another cause
  none kept    => there is a threshold, and this sweep brackets it

    python _pb_imgrowfit_gen.py gen
    python _pb_imgrowfit_gen.py pdf              # Word truth
    python _pb_imgrowfit_gen.py oxi              # Oxi, default
    python _pb_imgrowfit_gen.py oxi OXI_S1168=1  # Oxi, S1168 on
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_imgrowfit")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402
from _pb_cellimgtail_gen import (  # noqa: E402
    IMG_H_PT, IMG_W_PT, MARG, PGH, PGW, img_para, para, png_bytes,
)

PITCH = 13.8                 # measured in the sibling probe (TNR 12pt, single)
TOP = MARG / 20.0            # 72.0
BOT = (PGH - MARG) / 20.0    # 720.0
# Lines above the image. 1 = the 00161422 caption, 6 = the 002a301d empty run,
# 9 = the 0061c884 paragraph block -- the three real shapes that disagree.
BEFORES = [1, 2, 3, 6, 9]


def crossing_fill(k):
    """Smallest filler count whose row bottom passes the content bottom."""
    f = 0
    while TOP + PITCH * (1 + f) + k * PITCH + IMG_H_PT <= BOT:
        f += 1
    return f


def arms():
    out = []
    for k in BEFORES:
        c = crossing_fill(k)
        for step in (0, 2, 4, 6):
            out.append((k, c + step))
    return out


def docx():
    return os.path.join(OUT, "imgrowfit.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (k, fill) in enumerate(arms()):
        body.append(para("M%02d" % ai, pbb=ai > 0))
        for j in range(fill):
            body.append(para("a%df%d" % (ai, j)))
        # k lines, THEN the image last.
        cell = "".join(para("a%dB%d" % (ai, j + 1)) for j in range(k)) + img_para()
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
            + para("a%dside" % ai) + "</w:tc></w:tr></w:tbl>")
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
    print("wrote", docx(), len(arms()), "arms | image", IMG_H_PT, "pt | before", BEFORES)


def report(per, who):
    print("== %s ==  (content band %.1f .. %.1f)" % (who, TOP, BOT))
    print("%-4s %-4s %-6s %-9s %-22s %-9s %s"
          % ("arm", "k", "fill", "B1_y", "kept above image", "img page", "free@B1"))
    for ai, (k, fill) in enumerate(arms()):
        got = per.get(ai)
        if not got:
            print("%-4d %-4d %-6d MISSING" % (ai, k, fill))
            continue
        kept, imgpg, p0, rtop = got
        free = (BOT - rtop) if rtop else float("nan")
        print("%-4d %-4d %-6d %-9s %-22s %-9s %.1f"
              % (ai, k, fill, "%.1f" % rtop if rtop else "?",
                 " ".join(kept) if kept else "(none)",
                 "same" if imgpg == p0 else ("+%d" % (imgpg - p0)) if imgpg is not None else "?",
                 free))


def _collect(page_of, img_pages):
    per = {}
    for ai, (k, _f) in enumerate(arms()):
        start = page_of.get("M%02d" % ai)
        if not start:
            continue
        p0 = start[0]
        kept = [("B%d" % j) for j in range(1, k + 1)
                if page_of.get("a%dB%d" % (ai, j), (-1,))[0] == p0]
        b1 = page_of.get("a%dB1" % ai)
        rtop = b1[1] if b1 else None
        per[ai] = (kept, img_pages.get(ai), p0, rtop)
    return per


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
    page_of = {}
    for pi in range(doc.page_count):
        for bl in doc[pi].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t:
                    page_of.setdefault(t, (pi, round(ln["bbox"][1], 1)))
    img_pages = {}
    for ai, _a in enumerate(arms()):
        st = page_of.get("M%02d" % ai)
        if not st:
            continue
        for pi in (st[0], st[0] + 1, st[0] + 2):
            if pi >= doc.page_count:
                break
            if any(dr["rect"].height > 100 and abs(dr["rect"].width - IMG_W_PT) < 12
                   for dr in doc[pi].get_drawings()):
                img_pages[ai] = pi
                break
    report(_collect(page_of, img_pages), "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "imgrowfit_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "irf"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    page_of = {}
    img_on = {}
    for pi, pg in enumerate(pages):
        for e in pg["elements"]:
            t = (e.get("text") or "").strip()
            if t:
                page_of.setdefault(t, (pi, round(e["y"], 1)))
            if e.get("type") == "image" and (e.get("h") or 0) > 100:
                img_on.setdefault(pi, True)
    img_pages = {}
    for ai, _a in enumerate(arms()):
        st = page_of.get("M%02d" % ai)
        if not st:
            continue
        for pi in (st[0], st[0] + 1, st[0] + 2):
            if img_on.get(pi):
                img_pages[ai] = pi
                break
    report(_collect(page_of, img_pages), "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
