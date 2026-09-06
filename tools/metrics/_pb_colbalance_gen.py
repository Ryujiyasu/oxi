# -*- coding: utf-8 -*-
"""How does Word split the rows of a balanced two-column section whose rows
differ in height?  reference__0ea3ec86 p3: a 2-column continuous section holds
a heading (taller row) and 6 body lines; Word puts heading + 2 lines in column
1 and 4 lines in column 2 (max height 82.2), Oxi's S750 balances by ROW COUNT
(heading + 3 / 3, max 82.65).  Candidate laws: (a) fill column 1 while its
height stays <= total/2; (b) minimise the taller column.  They differ when the
heading is much taller than a body line, so the arms sweep the heading's
spacing-after.

    python _pb_colbalance_gen.py gen
    python _pb_colbalance_gen.py pdf      # Word COM export, then count lines per column
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_colbalance")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

# (label, heading spacing-after twips (None = no heading), heading sz half-points, n body lines)
ARMS = []
for n in (3, 4, 5, 6, 7, 8):
    ARMS.append(("h0_n%d" % n, None, 0, n))          # control: body lines only
    ARMS.append(("h200_n%d" % n, 200, 24, n))        # heading 12pt + 10pt after
    ARMS.append(("h600_n%d" % n, 600, 24, n))        # heading 12pt + 30pt after
    ARMS.append(("h1200_n%d" % n, 1200, 28, n))      # heading 14pt + 60pt after
for n in (1, 2, 3, 4):
    ARMS.append(("h200b_n%d" % n, 200, 21, n))       # same 28pt row, body-sized heading
    ARMS.append(("h360_n%d" % n, 360, 24, n))        # 18 + 18 = 36 (two grid lines)
    ARMS.append(("h100_n%d" % n, 100, 24, n))        # 18 + 5 = 23
    ARMS.append(("h800_n%d" % n, 800, 24, n))        # 18 + 40 = 58


def docx(label):
    return os.path.join(OUT, "colbalance_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    for label, hafter, hsz, n in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
                  '<w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
                  "<w:pPrDefault/></w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
        body = "<w:p><w:r><w:t>前のセクション</w:t></w:r></w:p>"
        body += ('<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>'
                 '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:pPr></w:p>')
        if hafter is not None:
            body += ('<w:p><w:pPr><w:spacing w:after="%d"/><w:rPr><w:sz w:val="%d"/></w:rPr></w:pPr>'
                     '<w:r><w:rPr><w:sz w:val="%d"/></w:rPr><w:t>見出し</w:t></w:r></w:p>' % (hafter, hsz, hsz))
        for i in range(n):
            body += "<w:p><w:r><w:t>本文%d行目</w:t></w:r></w:p>" % (i + 1)
        body += ('<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>'
                 '<w:cols w:num="2" w:space="425"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:pPr></w:p>')
        body += "<w:p><w:r><w:t>次のセクション</w:t></w:r></w:p>"
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>" + body
               + '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>'
                 '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
            z.writestr("word/styles.xml", styles)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def pdf():
    import fitz
    import win32com.client as w
    app = w.Dispatch("Word.Application")
    app.Visible = False
    try:
        for label, hafter, hsz, n in ARMS:
            out = docx(label)[:-5] + ".pdf"
            if not os.path.exists(out):
                d = app.Documents.Open(docx(label), ReadOnly=True)
                d.ExportAsFixedFormat(out, 17)
                d.Close(False)
            doc = fitz.open(out)
            page = doc[0]
            mid = page.rect.width / 2
            col = [[], []]
            heights = {}
            for b in page.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    t = "".join(s["text"] for s in l["spans"]).strip()
                    if not t or t in ("前のセクション", "次のセクション"):
                        if t == "次のセクション":
                            heights["next"] = l["bbox"][1]
                        continue
                    c = 0 if l["bbox"][0] < mid else 1
                    col[c].append((round(l["bbox"][1], 1), t))
            col[0].sort()
            col[1].sort()
            print("%-10s col1=%d %-30s col2=%d %-30s next_y=%.1f" % (
                label, len(col[0]), ",".join(t[:4] for _, t in col[0]), len(col[1]), ",".join(t[:4] for _, t in col[1]), heights.get("next", -1)))
            print("           y1=%s  y2=%s" % (",".join("%.1f" % y for y, _ in col[0]), ",".join("%.1f" % y for y, _ in col[1])))
    finally:
        app.Quit()


if __name__ == "__main__":
    {"gen": gen, "pdf": pdf}[sys.argv[1]]()
