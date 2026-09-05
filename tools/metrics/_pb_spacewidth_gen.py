# -*- coding: utf-8 -*-
"""How wide is an ASCII SPACE inside a proportional CJK face in Word?

educational__08709ff2 (BIZ UDPGothic 18pt, balanceSingleByteDoubleByteWidth,
compat 11): Word's PDF advances the space 9.0pt (= 0.5em) although the font's
own space glyph is 683/2048 = 0.333em (6.0pt, what Oxi uses); letters and
digits keep their proportional advances. One line gains 6pt over two spaces and
wraps a character that Oxi fits. Sweep face x the balance flag x compat and
read the space advance from Word's PDF export.

    python _pb_spacewidth_gen.py gen
    python _pb_spacewidth_gen.py pdf
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_spacewidth")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

FACES = ["BIZ UDPゴシック", "ＭＳ Ｐゴシック", "メイリオ", "游ゴシック", "ＭＳ ゴシック", "Century"]
ARMS = [("%s_%s_c%d" % (["bizudp", "mspg", "meiryo", "yugo", "msg", "century"][i], "bal" if bal else "nobal", compat), face, bal, compat)
        for i, face in enumerate(FACES) for bal in (False, True) for compat in (11, 15)]
TEXT = "国土 国土 A B 1 2 国"


def docx(label):
    return os.path.join(OUT, "spacewidth_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    for label, face, bal, compat in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
                  '<w:kern w:val="2"/><w:sz w:val="36"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
                  "<w:pPrDefault/></w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="left"/></w:pPr></w:style></w:styles>' % (face, face, face))
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                    '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>%s'
                    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/>'
                    "</w:compat></w:settings>" % ("<w:balanceSingleByteDoubleByteWidth/>" if bal else "", compat))
        body = '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % TEXT
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>'
                 '<w:docGrid w:type="lines" w:linePitch="576"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                       '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                       "</Relationships>")
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for label, *_ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                d.SaveAs2(docx(label)[:-5] + ".word.pdf", 17)
            finally:
                d.Close(False)
    finally:
        app.Quit()
    print("== WORD (PDF, 18pt): advance of each character of 「国土 国土 A B 1 2 国」 ==")
    for label, face, bal, compat in ARMS:
        pg = fitz.open(docx(label)[:-5] + ".word.pdf")[0]
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                chars = [c for sp in l["spans"] for c in sp["chars"]]
                t = "".join(c["c"] for c in chars)
                if t.startswith("国土"):
                    xs = [c["origin"][0] for c in chars]
                    advs = [(chars[i]["c"], round(xs[i + 1] - xs[i], 2)) for i in range(len(xs) - 1)]
                    sp = [a for c, a in advs if c == " "]
                    print("%-22s %-10s bal=%-5s c%d fonts=%s space=%s A=%s 1=%s" % (
                        label, face, bal, compat, sorted(set(s["font"] for s in l["spans"])), sp,
                        [a for c, a in advs if c == "A"], [a for c, a in advs if c == "1"]))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    else:
        gen()
