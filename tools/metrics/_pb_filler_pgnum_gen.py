# -*- coding: utf-8 -*-
"""Which number does Word give the blank filler page an oddPage / evenPage section inserts,
and where does the incoming section's pgNumType restart land?

reference__0ea3ec86: section 3 is oddPage with pgNumType start 91 after a page numbered
89; Word prints 91 on the first page after the blank. Oxi numbered the blank 91 and every
later page one too high.

Each arm: section 1 = cover pages (footer PAGE), section 2 = the type / restart under
test (same footer). Reads every page's footer number from the PDF ('-' = none printed).
One fresh Word process per arm.

    python tools/metrics/_pb_filler_pgnum_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_filler_pgnum")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
FOOTER = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr %s><w:p><w:r><w:t xml:space="preserve">PG=</w:t></w:r>'
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>') % W
# arm -> (cover pages, cover start, section-2 type, section-2 start or None, evenAndOddHeaders)
ARMS = {
    "odd_s91_after89": (2, 88, "oddPage", 91, False),
    "odd_s91_after89_eo": (2, 88, "oddPage", 91, True),
    "odd_none_after2": (2, None, "oddPage", None, False),
    "even_s10_after3": (3, None, "evenPage", 10, False),
    "odd_s5_after1": (1, None, "oddPage", 5, False),
}


def build(arm):
    n, st1, typ, st2, eo = ARMS[arm]
    fr = '<w:footerReference w:type="default" r:id="rId1"/>'
    pg = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    num1 = '<w:pgNumType w:start="%d"/>' % st1 if st1 else ""
    num2 = '<w:pgNumType w:start="%d"/>' % st2 if st2 else ""
    body = ""
    for i in range(n):
        body += '<w:p><w:r><w:t>COVER %d</w:t></w:r>%s</w:p>' % (i + 1, '<w:r><w:br w:type="page"/></w:r>' if i < n - 1 else "")
    body += '<w:p><w:pPr><w:sectPr>%s%s%s</w:sectPr></w:pPr></w:p>' % (fr, pg, num1)
    body += '<w:p><w:r><w:t>SECTION TWO</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/><w:t>SECTION TWO PAGE 2</w:t></w:r></w:p>'
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr>%s<w:type w:val="%s"/>%s%s</w:sectPr></w:body></w:document>' % (W, body, fr, typ, pg, num2))
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s>%s<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % (W, "<w:evenAndOddHeaders/>" if eo else "")
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "fp_%s.docx" % arm)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/footer1.xml", FOOTER)
        z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    return path


def footers(pdf):
    import pymupdf, re
    out = []
    for pg in pymupdf.open(pdf):
        t = pg.get_text()
        m = re.search(r"PG=(\d+)", t)
        body = "BLANK" if not re.search(r"COVER|SECTION", t) else re.search(r"(COVER \d|SECTION TWO PAGE 2|SECTION TWO)", t).group(1)
        out.append("%s:%s" % (body, m.group(1) if m else "-"))
    return out


if __name__ == "__main__":
    import win32com.client
    for arm in ARMS:
        p = build(arm); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        print("%-20s %s" % (arm, footers(pdf)), flush=True)
