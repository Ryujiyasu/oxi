# -*- coding: utf-8 -*-
"""Does an eastAsia face of «Times New Roman» take the break opportunities out of CJK text?

ja/correspondence__0d550693: «…取締役の David McMurtry 卿が永眠いたしました。…» -- Word
never lets «卿…» start on line 1, it moves the whole CJK run to line 2 and breaks it
there only at the margin (an emergency break of one long word).  On a faithful slice
the switch is the docDefaults eastAsia face: «Times New Roman» -> one word; Arial,
Arial MT, Courier New, Georgia, Calibri, Cambria, Palatino, «Times», Meiryo, ＭＳ 明朝 ->
ordinary CJK breaking; the eastAsia language does not matter.

Arms (synthetic, no styles part): eastAsia face TNR / ＭＳ 明朝 / Arial x lang ja-JP /
en-US x text with / without a Latin word before the CJK run.  Per arm: does line 1
end right after the Latin word (= CJK run moved whole)?

    python tools/metrics/_pb_ea_tnr_wordbreak_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_ea_tnr_wordbreak")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CJK = "卿が永眠いたしました。生前のご厚誼に深謝し、謹んでお知らせ申し上げます。"
TEXTS = {"latin": "共同創業者であり、非常勤取締役の David McMurtry " + CJK,
         "pure": "共同創業者であり、非常勤取締役のデイビッド・マクマートリー" + CJK}
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')


def build(name, ea, lang, text, cm=None):
    p = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="%s"/>'
         '<w:sz w:val="21"/><w:lang w:val="en-GB" w:eastAsia="%s"/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ea, lang, text))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1418" w:right="1418"/></w:sectPr>'
           '</w:body></w:document>' % (W, p))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, name + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        ct = CT
        if cm:
            ct = CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
            z.writestr("word/settings.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>' % (W, cm))
            z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", RELS); z.writestr("word/document.xml", doc)
    return path


def lines_of(docx):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    rows = {}
    for b in fitz.open(pdf)[0].get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip():
                        rows.setdefault(round(c["origin"][1]), []).append(c)
    return ["".join(c["c"] for c in sorted(v, key=lambda c: c["origin"][0])) for _, v in sorted(rows.items())]


def matrix():
    faces = os.environ.get("PB_FACES", "Times New Roman,Calibri,Cambria,Arial,Century,Courier New").split(",")
    for cm in (12, 14, 15):
        row = []
        for ea in faces:
            name = "m_%s_c%d" % (ea.replace(" ", ""), cm)
            # PB_TEXT2: another text; the verdict is then the end of line 1
            text = os.environ.get("PB_TEXT2", TEXTS["latin"])
            ls = lines_of(build(name, ea, "ja-JP", text, cm))
            row.append("%s:%s" % (ea, ("WORD" if ls[0].endswith("McMurtry") else "char") if text == TEXTS["latin"] else repr(ls[0][-6:])))
        print("compat %d  %s" % (cm, "  ".join(row)), flush=True)


if __name__ == "__main__" and os.environ.get("PB_MATRIX"):
    matrix()
elif __name__ == "__main__":
    for ea in ("Times New Roman", "ＭＳ 明朝", "Arial"):
        for lang in ("ja-JP", "en-US"):
            for tk, text in TEXTS.items():
                name = "%s_%s_%s" % (ea.replace(" ", ""), lang, tk)
                ls = lines_of(build(name, ea, lang, text))
                print("%-36s line1 ends %-16r line2 starts %r" % (name, ls[0][-10:], ls[1][:6] if len(ls) > 1 else None), flush=True)
