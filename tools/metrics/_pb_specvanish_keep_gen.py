# -*- coding: utf-8 -*-
"""Does a run-in heading (style separator: heading mark <w:vanish/><w:specVanish/>)
pass its keepNext on to the merged paragraph?

technical__0056b52f's glossary: every entry is a Heading3 (keepNext) whose mark is
hidden, running into a plain paragraph «– The element of …».  Oxi gives the merged
paragraph the heading's keepNext, so the whole glossary becomes one keepNext chain
and entries move to the next page that Word keeps (p20 «ES NUEMS body.»).

Arms: N filler lines, then [hidden-mark heading (keepNext) + continuation], then a
6-line keepLines paragraph that cannot fit.  heading keepNext on / off.  Per N: the
page of the merged paragraph.  If keepNext carries, the merged paragraph moves with
the keepLines block once that block does not fit.

    python tools/metrics/_pb_specvanish_keep_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_specvanish_keep")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
NS = list(range(44, 55))


def doc(n, keep):
    body = []
    for i in range(n):
        body.append('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>filler line %d</w:t></w:r></w:p>' % i)
    kn = "<w:keepNext/>" if keep else ""
    body.append('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/><w:rPr><w:vanish/><w:specVanish/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:b/></w:rPr><w:t>HEADWORD.</w:t></w:r></w:p>' % kn)
    body.append('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t xml:space="preserve"> – MERGED continuation text.</w:t></w:r></w:p>')
    big = " ".join(["block"] * 90)
    body.append('<w:p><w:pPr><w:keepLines/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>BLOCKSTART %s</w:t></w:r></w:p>' % big)
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>'
           '</w:body></w:document>' % (W, "".join(body)))
    return xml


CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')


def main():
    import fitz, win32com.client
    os.makedirs(OUT, exist_ok=True)
    for keep in (True, False):
        row = []
        for n in NS:
            p = os.path.join(OUT, "kn%d_n%d.docx" % (keep, n))
            with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/document.xml", doc(n, keep))
            pdf = p[:-5] + ".pdf"
            w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
            try:
                d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
            finally:
                w.Quit()
            pg_merged = pg_block = None
            for i, pg in enumerate(fitz.open(pdf)):
                t = pg.get_text()
                if "MERGED" in t and pg_merged is None:
                    pg_merged = i + 1
                if "BLOCKSTART" in t and pg_block is None:
                    pg_block = i + 1
            row.append("n%d:%s/%s" % (n, pg_merged, pg_block))
        print("heading keepNext=%s  (merged page / block page): %s" % (keep, " ".join(row)), flush=True)


if __name__ == "__main__":
    main()
