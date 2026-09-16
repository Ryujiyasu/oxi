# -*- coding: utf-8 -*-
"""On a linesAndChars grid, does Word balance two columns by grid SLOTS or by row HEIGHTS
when one row is a heading taller than the pitch?

reference__0ea3ec86 p36 (pitch 20.55, heading rows 30.75 = 1 pitch + spacing): the height
model gives Word's 27-row left column, the slot model (heading = 2 slots) 28.
reports__16785375 p10: the slot model gives Word's 10, the height model 11.

Sheet: continuous 2-column section on a linesAndChars grid (pitch 411 twips = 20.55pt),
N one-line body paragraphs (ＭＳ 明朝 11pt) and ONE heading paragraph (sz 22, spacing
after A twips, before B) at position K. Readout (Word PDF): how many rows land in column 1.

    python _pb_colgridbal_gen.py gen
    python _pb_colgridbal_gen.py pdf
"""
import os, sys, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_colgridbal")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

ARMS = []
for n in (12, 13, 14, 15, 16):
    for k in (0, 3, n // 2 - 1, n // 2, n // 2 + 1, n - 3):
        for a, b in ((200, 0), (0, 200), (120, 120), (0, 0)):
            ARMS.append(("n%d_k%d_a%d_b%d" % (n, k, a, b), n, k, a, b))

SECT1 = ('<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>'
         '<w:pgMar w:top="1304" w:right="1021" w:bottom="1134" w:left="1021" w:header="680" w:footer="567"/>'
         '<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="411" w:charSpace="2048"/></w:sectPr>')
SECT2 = SECT1.replace('<w:cols w:space="425"/>', '<w:cols w:num="2" w:space="425"/>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
          '<w:kern w:val="2"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
          "<w:pPrDefault/></w:docDefaults>"
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')


def docx(label):
    return os.path.join(OUT, "colgridbal_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    for label, n, k, a, b in ARMS:
        body = "<w:p><w:r><w:t>前のセクション</w:t></w:r></w:p>"
        body += "<w:p><w:pPr>" + SECT1 + "</w:pPr></w:p>"
        for i in range(n):
            if i == k:
                sp = '<w:spacing w:before="%d" w:after="%d"/>' % (b, a)
                body += ('<w:p><w:pPr>%s<w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>'
                         '<w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>見出し%d</w:t></w:r></w:p>' % (sp, i + 1))
            else:
                body += "<w:p><w:r><w:t>本文%d行目</w:t></w:r></w:p>" % (i + 1)
        body += "<w:p><w:pPr>" + SECT2 + "</w:pPr></w:p>"
        body += "<w:p><w:r><w:t>次のセクション</w:t></w:r></w:p>"
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>" + body + SECT1 + "</w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
            z.writestr("word/styles.xml", STYLES)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    try:
        for label, n, k, a, b in ARMS:
            out = docx(label)[:-5] + ".pdf"
            if not os.path.exists(out):
                d = app.Documents.Open(docx(label), ReadOnly=True)
                d.ExportAsFixedFormat(out, 17)
                d.Close(False)
            page = fitz.open(out)[0]
            mid = page.rect.width / 2
            col = [[], []]
            nxt = -1
            for blk in page.get_text("dict")["blocks"]:
                for l in blk.get("lines", []):
                    t = "".join(s["text"] for s in l["spans"]).strip()
                    if not t or t == "前のセクション":
                        continue
                    if t == "次のセクション":
                        nxt = l["bbox"][1]; continue
                    col[0 if l["bbox"][0] < mid else 1].append((round(l["bbox"][1], 1), t))
            col[0].sort(); col[1].sort()
            print("%-18s col1=%2d col2=%2d next_y=%.1f  y1=%s | y2=%s" % (
                label, len(col[0]), len(col[1]), nxt,
                ",".join("%.1f" % y for y, _ in col[0]), ",".join("%.1f" % y for y, _ in col[1])), flush=True)
    finally:
        app.Quit()


if __name__ == "__main__":
    {"gen": gen, "pdf": pdf}[sys.argv[1]]()
