# -*- coding: utf-8 -*-
"""A numbered paragraph whose STYLE (or pPr) gives leftChars while the
numbering level gives left / hanging in twips: where do the label and the text go?

`_pb_vert_list_gen.py` (List Paragraph leftChars 400 = 42pt, level 420/420 or
840/420): Word draws the label at 42 and the text at 63 -- as if the left were
leftChars + the level's hanging. Oxi takes the level's left (21 / 42).

    python _pb_list_leftchars_gen.py gen OUTDIR
    python _pb_list_leftchars_gen.py pdf OUTDIR   # Word: label x / text x / line-2 x from the margin
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

# (label, style pPr ind, level ind, direct ind)
ARMS = [
    ("s400_l420h420", '<w:ind w:leftChars="400" w:left="840"/>', 'w:left="420" w:hanging="420"', ""),
    ("s200_l420h420", '<w:ind w:leftChars="200" w:left="420"/>', 'w:left="420" w:hanging="420"', ""),
    ("s400_l420h210", '<w:ind w:leftChars="400" w:left="840"/>', 'w:left="420" w:hanging="210"', ""),
    ("s400_l1260h360", '<w:ind w:leftChars="400" w:left="840"/>', 'w:left="1260" w:hanging="360"', ""),
    ("stw840_l420h420", '<w:ind w:left="840"/>', 'w:left="420" w:hanging="420"', ""),
    ("direct400_l420h420", "", 'w:left="420" w:hanging="420"', '<w:ind w:leftChars="400" w:left="840"/>'),
    ("s400_l420h420_dh420", '<w:ind w:leftChars="400" w:left="840"/>', 'w:left="420" w:hanging="420"',
     '<w:ind w:hanging="420"/>'),
    ("s400_lvlhc200", '<w:ind w:leftChars="400" w:left="840"/>',
     'w:left="420" w:hangingChars="200" w:hanging="420"', ""),
]
TEXT = "見ゆ…（自然と）見える。もとは、マ上一「見」の未然形に自発・受身・可能の助動詞「ゆ」が付いて一語になった。助動詞「ゆ」は奈良時代ほどまで使われていた。"
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'


def docx(out, label):
    return os.path.join(out, "lc_%s.docx" % label)


def gen(out):
    os.makedirs(out, exist_ok=True)
    ct = CT.replace(
        "</Types>",
        '<Override PartName="/word/numbering.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
    rel = ('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
           'officeDocument/2006/relationships/%s" Target="%s"/>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             + rel % (1, "styles", "styles.xml") + rel % (2, "numbering", "numbering.xml")
             + "</Relationships>")
    for label, sind, lind, dind in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  "<w:docDefaults><w:rPrDefault>" + RPR + "</w:rPrDefault></w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:jc w:val="both"/></w:pPr></w:style>'
                  '<w:style w:type="paragraph" w:styleId="a8"><w:name w:val="List Paragraph"/>'
                  '<w:basedOn w:val="a"/><w:pPr>' + sind + "</w:pPr></w:style></w:styles>")
        num = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + ">"
               '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
               '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimalEnclosedCircle"/>'
               '<w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind ' + lind + "/></w:pPr></w:lvl>"
               '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
        item = ('<w:p><w:pPr><w:pStyle w:val="a8"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>'
                "</w:numPr>" + dind + "</w:pPr><w:r>" + RPR + "<w:t>" + TEXT + "</w:t></w:r></w:p>")
        plain = "<w:p><w:r>" + RPR + "<w:t>あいうえお本文の段落。</w:t></w:r></w:p>"
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + plain + item + item + plain
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1701" w:right="1701" w:bottom="1701" w:left="1701"/>'
               '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(docx(out, label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", num)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), out))


def rows(pdf_path):
    import fitz
    out = {}
    for b in fitz.open(pdf_path)[0].get_text("rawdict")["blocks"]:
        for line in b.get("lines", []):
            for span in line["spans"]:
                for c in span["chars"]:
                    if c["c"].strip():
                        out.setdefault(round(c["bbox"][3]), []).append((c["c"], round(c["bbox"][0] - 85.05, 2)))
    return [sorted(out[k], key=lambda t: t[1]) for k in sorted(out)]


def pdf(out):
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for arm in ARMS:
            src = docx(out, arm[0])
            dst = src.replace(".docx", ".pdf")
            d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.ExportAsFixedFormat(dst, 17)
                p = d.Paragraphs(2)
                com = (p.LeftIndent, p.FirstLineIndent, p.CharacterUnitLeftIndent,
                       p.CharacterUnitFirstLineIndent)
            finally:
                d.Close(False)
            r = rows(dst)
            print("%-22s %s | %s  COM L/F/Lch/Fch=%s" % (
                arm[0], " ".join("%s@%.2f" % c for c in r[1][:2]), "%s@%.2f" % r[2][0], com))
    finally:
        app.Quit()


if __name__ == "__main__":
    cmd, out = sys.argv[1], sys.argv[2]
    {"gen": gen, "pdf": pdf}[cmd](out)
