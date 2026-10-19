# -*- coding: utf-8 -*-
"""Does Word lay a glyph at its design advance or at the advance rounded to 10tw?

ja/educational__08709ff2 «…調べる方» (BIZ UDPGothic 18pt): the design widths sum
to 424.996 on a 425.2 line and Word keeps 方; Oxi rounds every advance to 10tw
(た 17.64 -> 17.5, ね 17.82 -> 18.0 ...) and wraps it.  Word's PDF origins step by
the design widths (た 17.64, し 15.34, す 17.14).

Arms: face x character, 20 copies of one character in a left-aligned paragraph;
the span from the first origin to the last divided by 19 = Word's advance.  Printed
next to the design advance and the 10tw-rounded one.

    python tools/metrics/_pb_advance_round_gen.py
"""
import glob, os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_advance_round")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
ARMS = [("BIZ UDPGothic", "たしね", 36), ("MS PGothic", "たしね", 36), ("MS PMincho", "たしね", 36),
        ("Meiryo", "たaW", 22), ("Yu Gothic", "たaW", 22), ("Calibri", "aeW", 22), ("Times New Roman", "aeW", 22),
        ("Arial", "aeW", 21), ("Century", "aeW", 21), ("BIZ UDPMincho", "たしね", 21)]
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')


def design(face, ch, hp):
    from fontTools.ttLib import TTFont, TTCollection
    for f in glob.glob(r"C:/Windows/Fonts/*.tt[cf]") + glob.glob(os.path.expandvars(r"%LOCALAPPDATA%/Microsoft/Windows/Fonts/*.tt[cf]")):
        try:
            fs = TTCollection(f).fonts if f.lower().endswith("ttc") else [TTFont(f)]
        except Exception:
            continue
        for t in fs:
            if face in {n.toUnicode() for n in t["name"].names if n.nameID in (1, 4)}:
                cm = t.getBestCmap()
                if ord(ch) in cm:
                    return t["hmtx"][cm[ord(ch)]][0] * hp / 2 / t["head"].unitsPerEm
    return None


def build():
    ps = []
    for face, chars, hp in ARMS:
        for ch in chars:
            ps.append('<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:kern w:val="0"/>'
                      '<w:sz w:val="%d"/></w:rPr><w:t>%s</w:t></w:r></w:p>' % (face, face, face, hp, ch * 20))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:bottom="1000" w:left="1000" w:right="1000"/></w:sectPr>'
           '</w:body></w:document>' % (W, "".join(ps)))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "advround.docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    import fitz, win32com.client
    docx = build(); pdf = docx[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    rows = []
    for pg in fitz.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                cs = sorted([c for s in l["spans"] for c in s["chars"]], key=lambda c: c["origin"][0])
                if len(cs) >= 20:
                    rows.append((cs[0]["c"], (cs[19]["origin"][0] - cs[0]["origin"][0]) / 19, l["spans"][0]["font"]))
    k = 0
    for face, chars, hp in ARMS:
        for ch in chars:
            c, adv, font = rows[k]; k += 1
            dz = design(face, ch, hp)
            r10 = None if dz is None else int(dz * 20 / 10 + 0.5) * 10 / 20
            print("%-16s %-3s %4.1fpt  word %.3f  design %s  10tw %s   (%s)" % (face, ch, hp / 2, adv, "%.3f" % dz if dz else "?", r10, font), flush=True)
