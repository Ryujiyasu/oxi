# -*- coding: utf-8 -*-
"""Inline pictures too wide to share a line: where does each one land?

zh dev technical__006f0290 has a paragraph holding two inline pictures
(343 x 228 / 343 x 224pt) on a 487pt measure: Word wraps the second under the
first (PDF 75..303, 311..535) and the next paragraph follows the second.  Oxi
gave each line the right height but advanced the paragraph by its text pitch,
so the second picture drew over the first.

Each arm is one page: marker «A», the subject paragraph (N pictures, 300pt wide
on a 468pt measure, so one per line), marker «B».  Read from Word's PDF: every
picture's top/bottom and the top of «B».

  python _pb_inline_pics_wrap_gen.py gen
  python _pb_inline_pics_wrap_gen.py word        # Word COM -> PDF -> truth json
  python _pb_inline_pics_wrap_gen.py oxi EXE     # Oxi --dump-layout, side by side
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_inline_pics_wrap")
EMU = 12700
PNG = (b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
       b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01"
       b"\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82")
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/img.png"/></Relationships>')

# (doc, east-asian font, docGrid xml)
DOCS = [
    ("nogrid", "Times New Roman", ""),
    ("grid", "SimSun", '<w:docGrid w:type="lines" w:linePitch="312"/>'),
]
HEIGHTS = [(228.0, 224.0), (100.0, 60.0), (30.0, 30.0), (100.0, 100.0, 40.0)]
LINES = [240, 360]


def styles(ea):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="%s" w:cs="Times New Roman"/>'
            '<w:sz w:val="21"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '</w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/></w:style></w:styles>' % ea)


def pic(idx, h):
    cx, cy = 300 * EMU, int(round(h * EMU))
    a = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="P%d"/>'
            '<a:graphic %s><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:nvPicPr><pic:cNvPr id="%d" name="img.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
            % (cx, cy, idx, idx, a, idx, cx, cy))


def arms():
    return [(line, hs) for line in LINES for hs in HEIGHTS]


def gen():
    os.makedirs(OUT, exist_ok=True)
    for name, ea, grid in DOCS:
        body, n = [], 1
        for i, (line, hs) in enumerate(arms()):
            sp = '<w:spacing w:line="%d" w:lineRule="auto"/>' % line
            body.append('<w:p><w:pPr>%s%s</w:pPr><w:r><w:t>A%d</w:t></w:r></w:p>'
                        % ("<w:pageBreakBefore/>" if i else "", sp, i))
            pics = "".join(pic(n + k, h) for k, h in enumerate(hs))
            n += len(hs)
            body.append('<w:p><w:pPr>%s</w:pPr>%s</w:p>' % (sp, pics))
            body.append('<w:p><w:pPr>%s</w:pPr><w:r><w:t>B%d</w:t></w:r></w:p>' % (sp, i))
        sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
                '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
                '%s</w:sectPr>' % grid)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>'
               + "".join(body) + sect + '</w:body></w:document>')
        p = os.path.join(OUT, name + ".docx")
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DRELS)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles(ea))
            z.writestr("word/media/img.png", PNG)
        print(p)


def word():
    import win32com.client
    import fitz
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    truth = {}
    try:
        for name, _, _ in DOCS:
            src = os.path.join(OUT, name + ".docx")
            pdf = os.path.join(OUT, name + "_word.pdf")
            d = app.Documents.Open(src, ReadOnly=True)
            d.ExportAsFixedFormat(pdf, 17)
            d.Close(False)
            truth[name] = read_pdf(fitz, pdf)
    finally:
        app.Quit()
    with open(os.path.join(OUT, "word.json"), "w", encoding="utf-8") as f:
        json.dump(truth, f, indent=1)
    show(truth)


def read_pdf(fitz, pdf):
    pages = []
    for pg in fitz.open(pdf):
        imgs = sorted(round(r["bbox"][1], 2) for r in pg.get_image_info())
        bots = sorted(round(r["bbox"][3], 2) for r in pg.get_image_info())
        b = [round(w[1], 2) for w in pg.get_text("words") if w[4].startswith("B")]
        pages.append({"img_top": imgs, "img_bot": bots, "B_top": b[:1]})
    return pages


def oxi(exe):
    res = {}
    for name, _, _ in DOCS:
        src = os.path.join(OUT, name + ".docx")
        dump = os.path.join(OUT, name + "_oxi.json")
        subprocess.run([exe, src, os.path.join(OUT, name + "_oxi"), "--dump-layout=" + dump],
                       capture_output=True)
        d = json.load(open(dump, encoding="utf-8"))
        pages = []
        for p in d["pages"]:
            ims = [e for e in p["elements"] if e["type"] == "image"]
            b = [e["y"] for e in p["elements"] if e["type"] == "text" and (e.get("text") or "").startswith("B")]
            pages.append({"img_top": sorted(round(e["y"], 2) for e in ims),
                          "img_bot": sorted(round(e["y"] + e["h"], 2) for e in ims), "B_top": b[:1]})
        res[name] = pages
    w = json.load(open(os.path.join(OUT, "word.json"), encoding="utf-8"))
    for name in res:
        for i, (line, hs) in enumerate(arms()):
            wp = w[name][i] if i < len(w[name]) else None
            op = res[name][i] if i < len(res[name]) else None
            print(f"{name} line={line} h={hs}\n   word {wp}\n   oxi  {op}")


def show(t):
    for name, pages in t.items():
        for i, (line, hs) in enumerate(arms()):
            print(name, line, hs, pages[i] if i < len(pages) else None)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    {"gen": gen, "word": word}.get(cmd, lambda: oxi(sys.argv[2]))()
