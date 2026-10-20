# -*- coding: utf-8 -*-
"""Which face does Word draw a missing font with, by the fontTable's declared charset?

ja/reference__2296511cc7b2fa07 declares «Noto Sans Japanese Bold» (not installed) with
charset 86 (GB2312) and a codePageRange that also has the Japanese bit; the new machine's
Word draws it in Microsoft YaHei, Oxi (S1403: Japanese code page -> Yu Gothic) in Yu Gothic.

Arms: a bogus face name declared with charset 80 / 86 / 88 / 81 / 00 / none, each with
and without the Japanese codePageRange bit; Japanese text 10.5pt.  One Word process per
arm (Word caches a missing face's substitute per process).  Per arm: the PDF span font.

    python tools/metrics/_pb_charset_substitute_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_charset_substitute")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FACE = "Oxiprobe Sans Missing"
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>')
ARMS = [(cs, jp) for cs in ("80", "86", "88", "81", "00", None) for jp in (False, True)]


def build(cs, jp):
    sig = '<w:sig w:usb0="20000003" w:usb1="2ADF3C10" w:usb2="00000016" w:usb3="00000000" w:csb0="%s" w:csb1="00000000"/>' % ("00020001" if jp else "00000001")
    font = '<w:font w:name="%s">%s<w:family w:val="auto"/><w:pitch w:val="default"/>%s</w:font>' % (
        FACE, '<w:charset w:val="%s"/>' % cs if cs else "", sig)
    ft = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts %s>%s</w:fonts>' % (W, font)
    run = ('<w:r><w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/></w:rPr>'
           '<w:t>もしも、共感できる生きテクがあったなら光栄です。</w:t></w:r>' % (FACE, FACE, FACE))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body><w:p>%s</w:p>'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>'
           '</w:body></w:document>' % (W, run))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "cs_%s_%s.docx" % (cs or "none", "jp" if jp else "nojp"))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/fontTable.xml", ft)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    import fitz, win32com.client
    for cs, jp in ARMS:
        docx = build(cs, jp); pdf = docx[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        fonts = sorted({s["font"] for b in fitz.open(pdf)[0].get_text("rawdict")["blocks"] for l in b.get("lines", []) for s in l["spans"]})
        print("charset %-4s jp-codepage %-5s -> %s" % (cs, jp, fonts), flush=True)
