# -*- coding: utf-8 -*-
"""Will Word start a Latin line with a mark that follows a space («rate %»)?

en/reports__001f1397, a Tahoma Bold 12 cell 80.55pt wide: «Actual recorded annual inflation
rate %» -> Word «inflation» / «rate %» although «inflation rate» (78.40) fits; Oxi «inflation
rate» / «%».  Hypothesis: the space before % is not a break opportunity.

Arms: «xxxx … word␠M tail» for marks M; the right indent is swept so the line can end
right after «word»; per mark, does Word ever put M first on the next line?

    python tools/metrics/_pb_lead_punct_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_lead_punct")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
MARKS = os.environ.get("PB_MARKS", "% ) , . : ; ! ? ] } $ ( - – — / ’ ” °").split(" ")
FACE = os.environ.get("PB_FACE", "Calibri")
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
TEXT = os.environ.get("PB_TEXT", "inflation rate %s tail words here")
RIGHTS = list(range(0, 1600, 40))
# PB_CELL=1: each paragraph inside a one-cell table whose width is swept instead
CELL = os.environ.get("PB_CELL") == "1"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;")


def build(mark):
    ps = []
    for k, r in enumerate(RIGHTS):
        run = ('<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="24"/><w:lang w:val="en-US"/></w:rPr>'
               '<w:t xml:space="preserve">%s</w:t></w:r>' % (FACE, FACE, esc(TEXT % mark)))
        if CELL:
            cw = 4000 - r
            ps.append('<w:p><w:pPr>%s<w:spacing w:after="0"/></w:pPr></w:p>' % ("<w:pageBreakBefore/>" if k else "")
                      + '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
                        '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>%s</w:p></w:tc></w:tr></w:tbl>'
                      % (cw, cw, cw, run))
        else:
            ps.append('<w:p><w:pPr>%s<w:spacing w:after="0"/><w:ind w:right="%d"/></w:pPr>%s</w:p>'
                      % ("<w:pageBreakBefore/>" if k else "", r, run))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="8000"/></w:sectPr>'
           '</w:body></w:document>' % (W, "".join(ps)))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "lp_%s_%x%s.docx" % (FACE.replace(" ", ""), ord(mark), "_cell" if CELL else ""))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/document.xml", doc)
    return path


def first_lines(pdf):
    import fitz
    out = []
    for pg in fitz.open(pdf):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        rows.setdefault(round(c["origin"][1]), []).append(c)
        out.append(["".join(c["c"] for c in sorted(v, key=lambda c: c["origin"][0])).strip() for _, v in sorted(rows.items())])
    return out


if __name__ == "__main__":
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for m in MARKS:
            docx = build(m); pdf = docx[:-5] + ".pdf"
            d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
            starts = set()
            ends = set()
            for ls in first_lines(pdf):
                for a, b in zip(ls, ls[1:]):
                    if b.startswith(m):
                        starts.add(a.split()[-1] + " | " + b.split()[0])
                    if a.endswith("rate"):
                        ends.add(a.split()[-1] + " | " + b.split()[0])
            print("%-2s  line starts with mark: %-24s  break after 'rate': %s" % (m, sorted(starts) or "never", sorted(ends) or "never"), flush=True)
    finally:
        w.Quit()
