# -*- coding: utf-8 -*-
"""Does an INLINE picture grow a line whose spacing rule is EXACT?

`creative__13152ea1` has a floating table whose first cell paragraph is
`<w:spacing w:line="320" w:lineRule="exact"/>` and holds a 256.9 x 120.75pt
inline drawing. Word's row is 130.8pt (the `trHeight` 2604tw floor + border):
the 16pt exact line does NOT grow for the picture, which simply overflows it.
Oxi's row is 147.75 -- it lets the picture set the line. The 17pt difference is
what turns the S1306 blank-line law's +10.25 into a page spill on that document.

Each arm is [基準] [a paragraph with a 100pt-tall inline picture under the rule
under test] [次] [末尾], and a second family puts the same paragraph in a
one-cell table with `trHeight` = 2604 (atLeast) and reads the paragraph after
the table. Word reports Info6 (collapsed starts); Oxi the dump's line y.

    python _pb_exactimg_gen.py gen
    python _pb_exactimg_gen.py pdf      # Word truth (COM Info6)
    python _pb_exactimg_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_exactimg")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import NS  # noqa: E402

MINCHO = "ＭＳ 明朝"
IMG_PT = 100.0                      # picture height in points (width 150)
EMU = 12700

# (label, in a table?, (rule, w:line) of the picture paragraph)
ARMS = [
    ("body_exact320", False, ("exact", 320)),
    ("body_exact240", False, ("exact", 240)),
    ("body_atleast320", False, ("atLeast", 320)),
    ("body_auto240", False, ("auto", 240)),
    ("body_noimg_exact320", None, ("exact", 320)),      # control: same line, no picture
    ("tbl_exact320", True, ("exact", 320)),
    ("tbl_atleast320", True, ("atLeast", 320)),
    ("tbl_auto240", True, ("auto", 240)),
]

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      "</Types>")
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>")
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
         '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
         "</Relationships>")
DNS = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
       'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
       'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def docx(label):
    return os.path.join(OUT, "exactimg_%s.docx" % label)


def png_bytes():
    import fitz
    pm = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 150, 100), 0)
    pm.clear_with(90)
    return pm.tobytes("png")


def picture():
    cx, cy = int(150 * EMU), int(IMG_PT * EMU)
    return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:docPr id="1" name="pic1"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="pic1"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' % (cx, cy, cx, cy))


def para(text, rule, line, with_pic=False):
    run = ('<w:r><w:t>%s</w:t></w:r>' % text) if text else ""
    if with_pic:
        run += picture()
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="%s"/></w:pPr>%s</w:p>'
            % (line, rule, run))


def table(inner):
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
            '<w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders>'
            '</w:tblPr><w:tblGrid><w:gridCol w:w="5394"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:val="2604"/></w:trPr>'
            '<w:tc><w:tcPr><w:tcW w:w="5394" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>' % inner)


def gen():
    os.makedirs(OUT, exist_ok=True)
    png = png_bytes()
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              '<w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="22"/>'
              "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
              '<w:jc w:val="both"/></w:pPr></w:style></w:styles>'
              % (MINCHO, MINCHO, MINCHO))
    for label, in_tbl, (rule, line) in ARMS:
        body = para("基準", "auto", 240)
        if in_tbl is None:
            body += para("画像", rule, line, with_pic=False)
        elif in_tbl:
            body += table(para("画像", rule, line, with_pic=True) + para("説明", "auto", 240))
        else:
            body += para("画像", rule, line, with_pic=True)
        body += para("次", "auto", 240) + para("末尾", "auto", 240)
        # NS may already declare some of these prefixes; a duplicate xmlns is malformed XML.
        extra = " ".join(a for a in DNS.split(" ") if a.split("=")[0] + "=" not in NS)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + " " + extra
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
                 '<w:docGrid w:type="linesAndChars" w:linePitch="375" w:charSpace="194"/>'
                 "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DRELS)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/media/image1.png", png)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def report(spans, who, extra=None):
    print("== %s ==" % who)
    print("%-22s %-6s %-11s %-10s %s" % ("arm", "table", "rule", "基準->次", "notes"))
    for label, in_tbl, (rule, line) in ARMS:
        sp = spans.get(label)
        print("%-22s %-6s %-11s %-10s %s"
              % (label, "-" if in_tbl is None else ("tbl" if in_tbl else "body"),
                 "%s%d" % (rule, line), "-" if sp is None else "%.2f" % sp,
                 (extra or {}).get(label, "")))


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    spans, extra = {}, {}
    try:
        for label, _, _ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                ys = {}
                for i in range(1, d.Paragraphs.Count + 1):
                    p = d.Paragraphs(i)
                    st = d.Range(p.Range.Start, p.Range.Start)
                    t = (p.Range.Text or "").rstrip("\r\x07")
                    ys.setdefault(t, float(st.Information(6)))
                if d.Tables.Count:
                    r = d.Tables(1).Range
                    extra[label] = "table rows(1).Height=%.2f" % d.Tables(1).Rows(1).Height
                spans[label] = ys["次"] - ys["基準"]
            finally:
                d.Close(False)
    finally:
        app.Quit()
    report(spans, "WORD (Info6, collapsed starts)", extra)


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    spans, extra = {}, {}
    for label, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "exactimg_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "ei"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        y, borders, by_y = {}, [], {}
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and (e.get("text") or "").strip():
                    # The dump splits a line into fragments: join them by y.
                    by_y.setdefault(round(e["y"], 2), []).append((e["x"], e["text"]))
                elif e["type"] == "border" and e.get("w", 0) > 100:
                    borders.append(e["y"])
        for yy, frags in sorted(by_y.items()):
            t = "".join(t for _, t in sorted(frags)).strip()
            for key in ("基準", "次"):
                if t.startswith(key):
                    y.setdefault(key, yy)
        if "基準" in y and "次" in y:
            spans[label] = y["次"] - y["基準"]
        if len(borders) >= 2:
            extra[label] = "table height=%.2f" % (max(borders) - min(borders))
    report(spans, "OXI " + (envs or "(default)"), extra)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
