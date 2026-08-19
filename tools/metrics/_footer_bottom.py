# -*- coding: utf-8 -*-
"""How far up does a tall footer push the body's last line?

technical__002c1ffa65f3a566 stops its page-6 body at y=572.84 although the
section's nominal content bottom is 629.35pt (841.95 page - 212.6 bottom
margin), leaving 56pt unused and refusing a 2-line entry that fits. Its footer
is eight 8pt paragraphs. Oxi fills to the nominal bottom and takes the entry,
which is the whole of that document's Phase-1 loss.

So sweep the one property: same page, same margins, same body, footer of N
paragraphs. The y of the LAST body line Word keeps on page 1 is the effective
content bottom, and how it moves with N is the rule.

    python _footer_bottom.py           # build, export through Word, measure
    python _footer_bottom.py --keep    # reuse the existing export
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_footer_bottom")

# the real section: A4, bottom margin 212.6pt, footer 170.1pt from the edge
PGH, BOTTOM_TW, FOOTER_TW = 16839, 4252, 3402
BODY_LINES = 60          # more than one page holds, so page 1 fills up
FOOTER_SZ = 16           # half-points -> 8pt, as in the document
ARMS = [0, 1, 2, 4, 6, 8, 12]

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
          '<w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
          '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
          '<w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')


def build(n_footer):
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "fb_%02d.docx" % n_footer)
    body = "".join(
        '<w:p><w:r><w:t xml:space="preserve">L%03d body line</w:t></w:r></w:p>' % i
        for i in range(BODY_LINES))
    sect = ('<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
            '<w:pgSz w:w="11907" w:h="%d"/>'
            '<w:pgMar w:top="1418" w:right="2410" w:bottom="%d" w:left="2410" '
            'w:header="720" w:footer="%d" w:gutter="0"/></w:sectPr>'
            % (PGH, BOTTOM_TW, FOOTER_TW))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<w:body>%s%s</w:body></w:document>' % (body, sect))
    fpar = ('<w:p><w:pPr><w:rPr><w:sz w:val="%d"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:sz w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">F%%02d footer</w:t></w:r></w:p>'
            % (FOOTER_SZ, FOOTER_SZ))
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '%s</w:ftr>' % "".join(fpar % i for i in range(n_footer)))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", DOCRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/footer1.xml", footer)
    return path


def export(docx, keep):
    pdf = docx[:-5] + ".pdf"
    if keep and os.path.exists(pdf):
        return pdf
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx, ReadOnly=True)
    try:
        d.ExportAsFixedFormat(pdf, 17)
    finally:
        d.Close(False)
        app.Quit()
    return pdf


def measure(pdf):
    """(last body line y on page 1, its text, footer top y)"""
    import fitz
    doc = fitz.open(pdf)
    body, foot = [], []
    for b in doc[0].get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            t = "".join(s["text"] for s in ln["spans"]).strip()
            if t.startswith("L"):
                body.append((round(ln["bbox"][1], 2), t))
            elif t.startswith("F"):
                foot.append(round(ln["bbox"][1], 2))
    body.sort()
    return (body[-1] if body else (None, "")), (min(foot) if foot else None), len(body)


def main():
    keep = "--keep" in sys.argv
    nominal = PGH / 20.0 - BOTTOM_TW / 20.0
    print("page %.2fpt, bottom margin %.2fpt, footer %.2fpt from edge"
          % (PGH / 20.0, BOTTOM_TW / 20.0, FOOTER_TW / 20.0))
    print("nominal content bottom = %.2fpt\n" % nominal)
    print("%-8s %-7s %-10s %-11s %-11s %s"
          % ("footer", "lines", "last_y", "footer_top", "eff_bottom", "vs nominal"))
    for n in ARMS:
        pdf = export(build(n), keep)
        (last_y, _t), ftop, nlines = measure(pdf)
        eff = last_y + 11.5 if last_y else 0.0     # + one line box
        print("%-8d %-7d %-10.2f %-11s %-11.2f %+.2f"
              % (n, nlines, last_y, "%.2f" % ftop if ftop else "-", eff, eff - nominal))


if __name__ == "__main__":
    main()
