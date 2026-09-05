# -*- coding: utf-8 -*-
"""Vertical (tbRl) per-character advance: when is it MORE than the font size?

educational__0828836's Word PDF walks its Meiryo columns at fs + 0.12pt per
character (10.5 bold -> 10.62, 9 regular -> 9.12; 600dpi-pixel alternation
88/89px), while _pb_vertgrid's plain arms walk at fs exactly. Sweep what the
document has and the probe lacked: compatibilityMode 15, useFELayout,
balanceSingleByteDoubleByteWidth, bold, kern, and a ruby paragraph.

    python _pb_vertadv_gen.py gen
    python _pb_vertadv_gen.py pdf      # Word truth (COM -> PDF glyph origins)
"""
import collections
import os
import statistics
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_vertadv")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

TEXT = "中納言参り給ひて御扇奉らせ給ふに隆家こそいみじき骨は得て侍れそれを張らせて参らせむ"
# (label, compat mode or None, extra compat elements, bold?, kern?, face)
ARMS = [
    ("plain_nosettings", None, "", False, True, "メイリオ"),
    ("c15", 15, "", False, True, "メイリオ"),
    ("c15_fe", 15, "<w:useFELayout/>", False, True, "メイリオ"),
    ("c15_bal", 15, "<w:balanceSingleByteDoubleByteWidth/>", False, True, "メイリオ"),
    ("c15_fe_bal", 15, "<w:useFELayout/><w:balanceSingleByteDoubleByteWidth/>", False, True, "メイリオ"),
    ("c15_bold", 15, "", True, True, "メイリオ"),
    ("c15_fe_bal_bold", 15, "<w:useFELayout/><w:balanceSingleByteDoubleByteWidth/>", True, True, "メイリオ"),
    ("c15_nokern", 15, "", False, False, "メイリオ"),
    ("c14_fe_bal", 14, "<w:useFELayout/><w:balanceSingleByteDoubleByteWidth/>", False, True, "メイリオ"),
    ("c15_fe_bal_msmincho", 15, "<w:useFELayout/><w:balanceSingleByteDoubleByteWidth/>", False, True, "ＭＳ 明朝"),
]


def docx(label):
    return os.path.join(OUT, "vertadv_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    for label, compat, extra, bold, kern, face in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
                  + ('<w:kern w:val="2"/>' if kern else "") +
                  '<w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
                  "<w:pPrDefault/></w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
        rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>%s<w:sz w:val="21"/></w:rPr>'
               % (face, face, face, "<w:b/><w:bCs/>" if bold else ""))
        body = "".join("<w:p><w:pPr>%s</w:pPr><w:r>%s<w:t>%s%d</w:t></w:r></w:p>" % (rpr, rpr, TEXT, i) for i in range(4))
        sect = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
                '<w:pgMar w:top="1701" w:right="1985" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>'
                '<w:textDirection w:val="tbRl"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>" + body + sect + "</w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct if compat is not None else CT)
            z.writestr("_rels/.rels", RELS)
            rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
            if compat is not None:
                rels += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                z.writestr("word/settings.xml",
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                           '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>' + extra +
                           '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/>' % compat +
                           "</w:compat></w:settings>")
            z.writestr("word/_rels/document.xml.rels", rels + "</Relationships>")
            z.writestr("word/styles.xml", styles)
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
    print("== WORD (PDF): mean per-character advance down the column (10.5pt) ==")
    for label, compat, extra, bold, kern, face in ARMS:
        pg = fitz.open(docx(label)[:-5] + ".word.pdf")[0]
        cols = collections.defaultdict(list)
        fonts = set()
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for sp in l["spans"]:
                    fonts.add(sp["font"])
                    for c in sp["chars"]:
                        if c["c"].strip():
                            cols[round(c["origin"][0])].append(c["origin"][1])
        advs = []
        for ys in cols.values():
            ys = sorted(ys)
            advs += [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
        print("%-22s compat=%-4s bold=%-5s kern=%-5s %s -> mean=%.4f min=%.2f max=%.2f n=%d fonts=%s"
              % (label, compat, bold, kern, face, statistics.mean(advs), min(advs), max(advs), len(advs), sorted(fonts)))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    else:
        gen()
