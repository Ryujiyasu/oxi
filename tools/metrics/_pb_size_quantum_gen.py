# -*- coding: utf-8 -*-
"""Does Word lay text out at a font size rounded to 1/600 inch (0.12pt)?

Every Word PDF on the truth machine writes span sizes on the 0.12pt lattice
(5 -> 5.04, 7 -> 6.96, 11 -> 11.04, 10 -> 9.96, 16 -> 15.96), and
ja/technical__c6d7cdb1's 5pt MS Mincho line advances 5.04pt per character -- one
character more than Oxi's 5.0 over a 33-character line.  This measures the
advance of runs of identical glyphs at several sizes:
  * ＭＳ 明朝 full-width 「あ」 x 40   (advance = em)
  * Arial 'x' x 60, Times New Roman 'n' x 60   (advance = width x size)
and compares the per-glyph advance with the nominal and the 0.12-rounded size.

    python tools/metrics/_pb_size_quantum_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_size_quantum")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
SIZES = [8, 9, 10, 11, 13, 14, 15, 21, 23]          # half-points: 4, 4.5, 5, 5.5, 6.5, 7, 7.5, 10.5, 11.5
CASES = [("ＭＳ 明朝", "あ" * 40), ("Arial", "x" * 60), ("Times New Roman", "n" * 60),
         ("Arial MT", "あ" * 40), ("Arial MT", "x" * 60)]   # Arial MT: a face missing on the machine
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')


def build():
    paras = []
    for face, text in CASES:
        for hp in SIZES:
            paras.append('<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr>'
                         '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/></w:rPr>'
                         '<w:t>%s</w:t></w:r></w:p>' % (face, face, face, hp, text))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="16838" w:h="11906"/><w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400"/></w:sectPr>'
           '</w:body></w:document>' % (W, "".join(paras)))
    os.makedirs(OUT, exist_ok=True)
    cm = os.environ.get("PB_COMPAT")   # e.g. 14: write a settings part with that compatibilityMode
    p = os.path.join(OUT, "sizes%s.docx" % (cm or ""))
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        ct = CT
        if cm:
            ct = CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
            z.writestr("word/settings.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%s"/></w:compat></w:settings>' % (W, cm))
            z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", RELS); z.writestr("word/document.xml", doc)
    return p


def main():
    import fitz, win32com.client
    p = build(); pdf = p[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    lines = []
    for pg in fitz.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [c for s in l["spans"] for c in s["chars"] if c["c"].strip()]
                if len(ch) > 10:
                    lines.append((l["spans"][0]["size"], l["spans"][0]["font"], ch))
    k = 0
    for face, text in CASES:
        for hp in SIZES:
            size, font, ch = lines[k]; k += 1
            adv = (ch[-1]["bbox"][0] - ch[0]["bbox"][0]) / (len(ch) - 1)
            nominal = hp / 2
            q = round(nominal / 0.12) * 0.12
            print("%-16s %5.2fpt  pdf size %.3f  advance %.4f  adv/nominal %.5f  adv/quantized %.5f" % (
                face, nominal, size, adv, adv / nominal, adv / q))


if __name__ == "__main__":
    main()
