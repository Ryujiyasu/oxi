# -*- coding: utf-8 -*-
"""Is a paragraph's spacing-before kept at the top of a page when the page
starts with a SECTION BREAK (nextPage / oddPage / evenPage) or a manual page
break, as opposed to a natural flow break?

reference__0ea3ec86 p3: the heading 1 「❖ 障害者に関するマーク」 (before 240,
exact 360, after 240) opens an oddPage section; Word's PDF puts its glyph top
at 79.5 = margin 65.2 + 12 + ~2, i.e. the 12pt before is KEPT; Oxi's page-top
suppression drops it and the column-balance of the section below shifts.

    python _pb_secbefore_gen.py gen
    python _pb_secbefore_gen.py com
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_secbefore")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

SECT = ('<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1304" w:right="1021" w:bottom="1134" w:left="1021" '
        'w:header="680" w:footer="567"/><w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="411"/>')
# (label, how the 2nd page starts)
ARMS = [("nextPage", "sect:nextPage"), ("oddPage", "sect:oddPage"), ("evenPage", "sect:evenPage"),
        ("pagebreak", "br"), ("flow", "flow"), ("continuous", "sect:continuous")]
ARMS += [(l + "_c15", h + "|c15") for l, h in ARMS]
ARMS += [(l + "_c14", h + "|c14") for l, h in ARMS[:6]]
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + '><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%s"/>'
            '</w:compat></w:settings>')
FILL = 33  # lines that fill page 1 (band 65.2..785.2 = 720pt / 20.55 = 35 lines)


def docx(label):
    return os.path.join(OUT, "secbefore_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    for label, how in ARMS:
        compat = how.split("|")[1] if "|" in how else None
        how = how.split("|")[0]
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
                  '<w:kern w:val="2"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
                  "<w:pPrDefault/></w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
        n = FILL if how != "flow" else 35
        body = "".join("<w:p><w:r><w:t>埋%d</w:t></w:r></w:p>" % i for i in range(n))
        if how.startswith("sect:"):
            kind = how.split(":")[1]
            body += '<w:p><w:pPr><w:sectPr><w:type w:val="%s"/>%s</w:sectPr></w:pPr></w:p>' % (kind, SECT)
        elif how == "br":
            body += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
        body += ('<w:p><w:pPr><w:spacing w:before="240" w:after="240" w:line="360" w:lineRule="exact"/>'
                 '<w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>見出し</w:t></w:r></w:p>')
        body += "<w:p><w:r><w:t>本文の一行目</w:t></w:r></w:p>"
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>" + body
               + "<w:sectPr>" + SECT + "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            ct = CT
            if compat:
                ct = CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                       '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
            z.writestr("word/styles.xml", styles)
            z.writestr("word/document.xml", doc)
            if compat:
                z.writestr("word/settings.xml", SETTINGS % compat[1:])
    print("wrote %d arms" % len(ARMS))


def com():
    import win32com.client as w
    app = w.Dispatch("Word.Application")
    app.Visible = False
    try:
        for label, how in ARMS:
            if len(sys.argv) > 2 and sys.argv[2] not in label:
                continue
            d = app.Documents.Open(docx(label), ReadOnly=True)
            out = []
            for i in range(1, d.Paragraphs.Count + 1):
                p = d.Paragraphs(i)
                t = p.Range.Text.strip()
                if t.startswith("見出し") or t.startswith("本文"):
                    r = d.Range(p.Range.Start, p.Range.Start)
                    out.append("%s pg=%d y=%.2f" % (t[:3], r.Information(3), r.Information(6)))
            print("%-11s %s" % (label, " | ".join(out)))
            d.Close(False)
    finally:
        app.Quit()


if __name__ == "__main__":
    {"gen": gen, "com": com}[sys.argv[1]]()
