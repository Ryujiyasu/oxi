# -*- coding: utf-8 -*-
"""WHERE does Word paint an inline picture that is taller than its EXACT line?

S1320 settled the flow (an exact line does not grow for its inline object);
legal__02f84965's page 8 shows the paint side: Word draws the 113.85pt box so
that its bottom lands at 769 (the content bottom is 771) although its host line
sits at 720 -- Oxi paints from the host line's top and runs 55pt past the
margin. Sweep the host line's distance to the page bottom and read, from Word's
own PDF export, the picture rectangle against the host line's text baseline.

    python _pb_exactimg_pos_gen.py gen
    python _pb_exactimg_pos_gen.py pdf      # Word truth (COM -> PDF: image rect + text rows)
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_exactimg_gen as P  # noqa: E402
from _pb_pxgrid_gen import NS  # noqa: E402

OUT = os.path.join(P.REPO, "pipeline_data", "_pb_exactimg_pos")
# (label, number of 12pt filler lines before the host, rule, w:line)
# content height = 842 - 2 x 56.7 = 728.6pt; a 12pt auto line is ~15.6pt (grid 18.75)
ARMS = [
    ("exact_mid", 6, "exact", 320),
    ("exact_room100", 28, "exact", 320),     # ~ 525pt down: 100pt still fits below
    ("exact_room60", 31, "exact", 320),      # ~ 580pt: 60pt of room
    ("exact_room20", 33, "exact", 320),      # ~ 620pt: 20pt of room
    ("exact_room0", 35, "exact", 320),       # the host itself is the last line that fits
    ("auto_mid", 6, "auto", 240),
    ("auto_room20", 33, "auto", 240),
]


def docx(label):
    return os.path.join(OUT, "exactimgpos_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    png = P.png_bytes()
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
              % (P.MINCHO, P.MINCHO, P.MINCHO))
    for label, nfill, rule, line in ARMS:
        body = "".join(P.para("基準%d" % i, "auto", 240) for i in range(nfill))
        body += P.para("画像", rule, line, with_pic=True)
        body += P.para("次", "auto", 240) + P.para("末尾", "auto", 240)
        extra = " ".join(a for a in P.DNS.split(" ") if a.split("=")[0] + "=" not in NS)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + " " + extra
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
                 '<w:docGrid w:type="linesAndChars" w:linePitch="375" w:charSpace="194"/>'
                 "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", P.CT)
            z.writestr("_rels/.rels", P.RELS)
            z.writestr("word/_rels/document.xml.rels", P.DRELS)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/media/image1.png", png)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def measure(path):
    import fitz
    doc = fitz.open(path)
    out = []
    for pno, pg in enumerate(doc, 1):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                chars = [c for sp in l["spans"] for c in sp["chars"] if c["c"].strip()]
                if chars:
                    rows["".join(c["c"] for c in chars)[:4]] = round(chars[0]["origin"][1], 2)
        imgs = [tuple(round(v, 2) for v in r) for info in pg.get_image_info() for r in [info["bbox"]]]
        out.append((pno, rows.get("画像"), rows.get("次"), rows.get("末尾"), imgs))
    return out


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for label, *_ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                d.SaveAs2(docx(label)[:-5] + ".word.pdf", 17)
            finally:
                d.Close(False)
    finally:
        app.Quit()
    print("== WORD (PDF): per page -> host baseline 画像, next 次, 末尾, image bbox (x0,y0,x1,y1); page bottom = 785.3 ==")
    for label, nfill, rule, line in ARMS:
        for pno, host, nxt, tail, imgs in measure(docx(label)[:-5] + ".word.pdf"):
            print("%-14s %-5s p%d host=%s next=%s tail=%s img=%s" % (label, rule, pno, host, nxt, tail, imgs))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    else:
        gen()
