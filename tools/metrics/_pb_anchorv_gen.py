# -*- coding: utf-8 -*-
"""Where does Word measure `positionV relativeFrom="paragraph"` from?

educational__0214ac95 -- the document holding the CJK font batch -- is a
worksheet whose 268 of 359 paragraphs live in 142 text boxes, anchored by 79
`wp:anchor` elements that all use `relativeFrom="paragraph"` with an offset.
Word fits it on two pages with 38 lines on the first; Oxi puts 21 lines there
and pushes elements to y=693/751/945 against a page bottom of 841.95, then
spills onto a third page.  Every line-height path involved has now been shown to
match Word (auto, atLeast, `line=0 atLeast`, grid cells), so the remaining
suspect is the anchor's reference point.

Each arm puts three marker paragraphs on a page and hangs one anchored text box
off the middle one, varying only what the offset is measured from and how big it
is.  Reading the box's own text against the markers gives the reference point
directly: the anchoring paragraph's TOP, its bottom, or the cursor after it.

  python _pb_anchorv_gen.py gen
  python _pb_anchorv_gen.py pdf      # Word truth
  python _pb_anchorv_gen.py oxi      # Oxi, same arms
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_anchorv")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

EMU = 12700
WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'

# (arm, relativeFrom, offset pt, anchor paragraph index within the page,
#  body line size half-points)
ARMS = [
    ("para_off0", "paragraph", 0.0, 1, 24),
    ("para_off20", "paragraph", 20.0, 1, 24),
    ("para_off0_p2", "paragraph", 0.0, 2, 24),   # anchor on the 3rd paragraph
    ("para_big_line", "paragraph", 0.0, 1, 48),  # 24pt body: does the ref move?
    ("line_off0", "line", 0.0, 1, 24),
    ("margin_off0", "margin", 0.0, 1, 24),
    ("page_off40", "page", 40.0, 1, 24),
]
BOX_W, BOX_H = 120.0, 24.0


def docx():
    return os.path.join(OUT, "anchorv.docx")


def marker(tag, ai, brk):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"'
            ' w:hAnsi="Arial"/><w:sz w:val="14"/></w:rPr><w:t>%s%02dZ</w:t>'
            "</w:r></w:p>" % ("<w:pageBreakBefore/>" if brk else "", tag, ai))


def anchored(pid, rel, off_pt, sz_hp):
    cx, cy = int(BOX_W * EMU), int(BOX_H * EMU)
    off = int(off_pt * EMU)
    rpr = ('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="18"/>'
           "</w:rPr>")
    inner = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
             ' w:lineRule="auto"/></w:pPr><w:r>' + rpr + "<w:t>BOX</w:t></w:r></w:p>")
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
        ' w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"'
        ' w:hAnsi="Times New Roman"/><w:sz w:val="%d"/></w:rPr>'
        "<w:t>anchor host line</w:t></w:r>" % sz_hp
        + f'<w:r><w:drawing><wp:anchor {WP_NS} distT="0" distB="0" distL="114300" '
        'distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" '
        'layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>2540000</wp:posOffset>'
        "</wp:positionH>"
        f'<wp:positionV relativeFrom="{rel}"><wp:posOffset>{off}</wp:posOffset>'
        "</wp:positionV>"
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        "<wp:wrapNone/>"
        f'<wp:docPr id="{pid}" name="AV{pid}"/><wp:cNvGraphicFramePr/>'
        f"<a:graphic {A_NS}>"
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f"<wps:wsp {WPS_NS}><wps:cNvSpPr/><wps:spPr>"
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
        "</wps:spPr>"
        f"<wps:txbx><w:txbxContent>{inner}</w:txbxContent></wps:txbx>"
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" '
        'bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>")


def body_line(sz_hp, n):
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"'
            ' w:hAnsi="Times New Roman"/><w:sz w:val="%d"/></w:rPr>'
            "<w:t>host paragraph %d</w:t></w:r></w:p>" % (sz_hp, n))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (name, rel, off, host_idx, sz) in enumerate(ARMS):
        body.append(marker("A", ai, ai > 0))
        for k in range(3):
            if k == host_idx:
                body.append(anchored(100 + ai, rel, off, sz))
            else:
                body.append(body_line(sz, k))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11907" w:h="16839"/>'
           '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:eastAsia="MS Mincho"'
              ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="21"/></w:rPr></w:style>'
              "</w:styles>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms")


def report(per, who):
    print("== %s ==" % who)
    print("%-14s %-10s %6s %9s %9s %9s"
          % ("arm", "relFrom", "off", "host_top", "box_top", "box-host"))
    for ai, (name, rel, off, host_idx, sz) in enumerate(ARMS):
        g = per.get(ai) or {}
        h, b = g.get("host"), g.get("box")
        if h is None or b is None:
            print("%-14s %-10s %6.1f   MISSING (host=%s box=%s)" % (name, rel, off, h, b))
            continue
        print("%-14s %-10s %6.1f %9.2f %9.2f %9.2f" % (name, rel, off, h, b, b - h))


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
        for m in re.finditer(r"A(\d\d)Z", doc[pi].get_text()):
            page_of.setdefault(int(m.group(1)), pi)
    per = {}
    for ai in range(len(ARMS)):
        pi = page_of.get(ai)
        if pi is None:
            continue
        g = {}
        for bl in doc[pi].get_text("dict")["blocks"]:
            for ln in bl.get("lines", []):
                for sp in ln["spans"]:
                    t = sp["text"].strip()
                    top = round(sp["origin"][1] - sp["ascender"] * sp["size"], 2)
                    if t == "BOX":
                        g["box"] = top
                    elif t.startswith("anchor host"):
                        g["host"] = top
        per[ai] = g
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "anchorv_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "av"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    page_of = {}
    for pi, pg in enumerate(pages):
        for e in pg["elements"]:
            m = re.fullmatch(r"A(\d\d)Z", (e.get("text") or "").strip())
            if m:
                page_of.setdefault(int(m.group(1)), pi)
    per = {}
    for ai in range(len(ARMS)):
        pi = page_of.get(ai)
        if pi is None:
            continue
        g = {}
        for e in pages[pi]["elements"]:
            t = (e.get("text") or "").strip()
            if t == "BOX":
                g["box"] = round(e["y"], 2)
            elif t.startswith("anchor") or t.startswith("host paragraph"):
                g.setdefault("host", round(e["y"], 2))
        per[ai] = g
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
