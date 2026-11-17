# -*- coding: utf-8 -*-
"""Numbered / bulleted paragraphs in a tbRl section: where does Word put the
label, and where does the text after it start, down the column?

ja/educational__047ff775 (decimalEnclosedCircle, level left 420 hanging 420) and
ja/creative__9c1d5eb7 (bullet ＊, level left 1076 hanging 360, direct leftChars=0)
draw the label at left - hanging and the text at left. These arms vary the
label width against the hanging, the suffix and the hanging itself.

    python _pb_vert_list_gen.py gen OUTDIR
    python _pb_vert_list_gen.py pdf OUTDIR    # Word truth: glyph tops per column
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

# (label, numFmt, lvlText, level left, level hanging, suff, direct ind)
ARMS = [
    ("circle_420_420", "decimalEnclosedCircle", "%1", 420, 420, None, ""),
    ("bullet_1076_360_lc0", "bullet", "＊", 1076, 360, None, '<w:ind w:leftChars="0"/>'),
    ("decimal_840_420", "decimal", "%1.", 840, 420, None, ""),
    ("wide_420_420", "decimal", "第%1条", 420, 420, None, ""),
    ("circle_space", "decimalEnclosedCircle", "%1", 420, 420, "space", ""),
    ("circle_nothing", "decimalEnclosedCircle", "%1", 420, 420, "nothing", ""),
    ("circle_nohang", "decimalEnclosedCircle", "%1", 0, 0, None, ""),
    ("fwdigit_840_420", "decimalFullWidth", "（%1）", 840, 420, None, ""),
]
TEXT = "見ゆ…（自然と）見える。もとは、マ上一「見」の未然形に自発・受身・可能の助動詞「ゆ」が付いて一語になった。助動詞「ゆ」は奈良時代ほどまで使われていた。"
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'


def docx(out, label):
    return os.path.join(out, "vlist_%s.docx" % label)


def numbering(fmt, text, left, hang, suff):
    s = '<w:suff w:val="%s"/>' % suff if suff else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + ">"
            '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="%s"/>%s<w:lvlText w:val="%s"/>'
            '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="%d" w:hanging="%d"/></w:pPr></w:lvl>'
            '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
            % (fmt, s, text, left, hang))


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
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault>" + RPR + "</w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:pPr><w:jc w:val="both"/></w:pPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="a8"><w:name w:val="List Paragraph"/>'
              '<w:basedOn w:val="a"/><w:pPr><w:ind w:leftChars="400" w:left="840"/></w:pPr></w:style>'
              "</w:styles>")
    for label, fmt, text, left, hang, suff, ind in ARMS:
        item = ('<w:p><w:pPr><w:pStyle w:val="a8"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>'
                "</w:numPr>" + ind + "</w:pPr><w:r>" + RPR + "<w:t>" + TEXT + "</w:t></w:r></w:p>")
        plain = "<w:p><w:r>" + RPR + "<w:t>あいうえお本文の段落。</w:t></w:r></w:p>"
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + plain + item + item + plain
               + '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
               '<w:pgMar w:top="1701" w:right="1985" w:bottom="1701" w:left="1701"/>'
               '<w:textDirection w:val="tbRl"/><w:docGrid w:type="lines" w:linePitch="360"/>'
               "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(out, label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", numbering(fmt, text, left, hang, suff))
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), out))


def columns(pdf_path):
    """[(x_centre, [(char, top), ...]), ...] right to left."""
    import fitz
    cols = {}
    for b in fitz.open(pdf_path)[0].get_text("rawdict")["blocks"]:
        for line in b.get("lines", []):
            for span in line["spans"]:
                for c in span["chars"]:
                    if not c["c"].strip():
                        continue
                    x = round((c["bbox"][0] + c["bbox"][2]) / 2, 1)
                    key = next((k for k in cols if abs(k - x) < 4), x)
                    cols.setdefault(key, []).append((c["c"], round(c["bbox"][1], 2)))
    return [(k, sorted(cols[k], key=lambda t: t[1])) for k in sorted(cols, reverse=True)]


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
            finally:
                d.Close(False)
            print(arm[0])
            for x, cs in columns(dst)[:4]:
                print("   x=%.1f %s" % (x, " ".join("%s@%.2f" % c for c in cs[:5])))
    finally:
        app.Quit()


if __name__ == "__main__":
    cmd, out = sys.argv[1], sys.argv[2]
    {"gen": gen, "pdf": pdf}[cmd](out)
