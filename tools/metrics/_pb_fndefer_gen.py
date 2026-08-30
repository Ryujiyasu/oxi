# -*- coding: utf-8 -*-
"""When a footnote's text will not fit, does Word stop the BODY at its reference?

`reports__0018715b4769984f` p5 (EN Phase-1 0.9560): Word ends the body at 639.22
and leaves 64pt -- 4.4 body lines -- empty above the footnote separator, then
starts p6 with a plain 3-line paragraph that carries no keepNext, no keepLines
and would fit in that gap. The page's last body line carries footnote reference
17, and 17's TEXT is not on that page: p5 shows footnotes 14/15/16 and p6 opens
with 17.

So the reference stayed and the note moved -- and the body stopped there. This
probe asks whether stopping is the rule.

Each arm is one page of single-line filler paragraphs. Line `REF` carries a
footnote whose text length is swept, and more filler follows. If Word stops the
body at the reference when the note cannot fit, the last body line on page 1 is
REF and the rest of the column is blank. If it only defers the note, the body
keeps filling to the bottom.

    python tools/metrics/_pb_fndefer_gen.py         # build
    python tools/metrics/_pb_fndefer_read.py word   # Word truth (COM -> PDF)
    python tools/metrics/_pb_fndefer_read.py oxi    # Oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_fndefer"
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"""
SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr>
<w:compat><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

SEP = ('<w:footnote w:type="separator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" '
       'w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
       '<w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:pPr><w:spacing '
       'w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/>'
       "</w:r></w:p></w:footnote>")

# The reference sits this many single-line paragraphs down the page; the footnote
# text is `fnlen` sentences long (each ~1 footnote line at 10pt across the column).
REFS = [38, 40, 42]
FNLENS = [1, 3, 6]
SENT = ("This is footnote sentence number %d and it is written to be long enough "
        "that it occupies a full line of the footnote area on its own. ")

ARMS = [("ref%d_fn%d" % (r, f), r, f) for r in REFS for f in FNLENS]


def build(tag, nref, fnlen):
    fn_text = "".join(SENT % (i + 1) for i in range(fnlen))
    body = ""
    for i in range(nref):
        body += ('<w:p><w:r><w:t xml:space="preserve">BODY%03d filler line</w:t>'
                 "</w:r></w:p>" % (i + 1))
    body += ('<w:p><w:r><w:t xml:space="preserve">REFLINE carries the note</w:t></w:r>'
             '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
             '<w:footnoteReference w:id="2"/></w:r></w:p>')
    for i in range(14):
        body += ('<w:p><w:r><w:t xml:space="preserve">AFTER%03d filler line</w:t>'
                 "</w:r></w:p>" % (i + 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>\n" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
           ' w:header="708" w:footer="708" w:gutter="0"/>'
           "</w:sectPr></w:body></w:document>")
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                 + SEP +
                 '<w:footnote w:id="2"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
                 'w:lineRule="auto"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>'
                 '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
                 '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:footnote>'
                 "</w:footnotes>" % fn_text)
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/footnotes.xml", footnotes)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    for a in ARMS:
        build(*a)
    print("built %d arms in %s" % (len(ARMS), OUT))
