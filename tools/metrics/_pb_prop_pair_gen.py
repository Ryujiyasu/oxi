# -*- coding: utf-8 -*-
"""Does Word halve a mark of a yakumono PAIR when the font's mark is already narrow?

ja/educational__08709ff2: BIZ UDPGothic 18pt, compressPunctuation, compat 11 --
«このうち、「重ねる…» puts 、 at its natural 10.80 (0.6 em, hmtx) and 「 at 9.0,
where Oxi halves the 、 of the 、「 pair (S1494/S1217 rule) to 5.5 and fits one more
character per line.  A full-width font (MS Mincho, technical__c5bb0090) does halve.

Arms: eastAsia face x compat 11/15 x characterSpacingControl compressPunctuation /
doNotCompress; one left-aligned paragraph per pair text.  Per arm: the advance of
the first mark of the pair next to its hmtx width.

    python tools/metrics/_pb_prop_pair_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_prop_pair")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FACES = os.environ.get("PB_FACES", "BIZ UDPGothic,BIZ UDGothic,MS PGothic,MS Gothic,Meiryo,Yu Gothic").split(",")
TEXTS = ["甲亜、「亜亜", "甲亜」「亜亜", "甲亜」、亜亜", "甲亜（（亜亜", "甲亜。」亜亜", "甲亜、亜亜亜"]
SZ = 36
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')


def build(face, cm, csc):
    ps = "".join(
        '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
        '<w:sz w:val="%d"/><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>%s</w:t></w:r></w:p>' % (face, face, face, SZ, t) for t in TEXTS)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1418" w:right="1418"/></w:sectPr>'
           '</w:body></w:document>' % (W, ps))
    st = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="%s"/><w:compat>'
          '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>' % (W, csc, cm))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "%s_c%d_%s.docx" % (face.replace(" ", ""), cm, csc[:4]))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/settings.xml", st)
        z.writestr("word/document.xml", doc)
    return path


def rows_of(pdf):
    import fitz
    out = []
    for b in fitz.open(pdf)[0].get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            ch = sorted([c for s in l["spans"] for c in s["chars"]], key=lambda c: c["origin"][0])
            if ch and ch[0]["c"] == "甲":
                out.append(ch)
    return out


def hmtx(face, ch):
    from fontTools.ttLib import TTFont, TTCollection
    import glob
    for f in glob.glob(r"C:/Windows/Fonts/*.tt[cf]") + glob.glob(os.path.expandvars(r"%LOCALAPPDATA%/Microsoft/Windows/Fonts/*.tt[cf]")):
        try:
            fs = TTCollection(f).fonts if f.lower().endswith("ttc") else [TTFont(f)]
        except Exception:
            continue
        for t in fs:
            names = {n.toUnicode() for n in t["name"].names if n.nameID in (1, 4)}
            if face in names:
                return round(t["hmtx"][t.getBestCmap()[ord(ch)]][0] * SZ / 2 / t["head"].unitsPerEm, 2)
    return None


if __name__ == "__main__":
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for face in FACES:
            print("== %s  hmtx 、=%s 「=%s 」=%s （=%s 。=%s" % (face, *(hmtx(face, c) for c in "、「」（。")), flush=True)
            for cm in (11, 15):
                for csc in ("compressPunctuation", "doNotCompress"):
                    docx = build(face, cm, csc); pdf = docx[:-5] + ".pdf"
                    d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
                    adv = ["%s%.2f" % (r[2]["c"], r[3]["origin"][0] - r[2]["origin"][0]) for r in rows_of(pdf)]
                    print("   c%d %-20s %s" % (cm, csc, "  ".join(adv)), flush=True)
    finally:
        w.Quit()
