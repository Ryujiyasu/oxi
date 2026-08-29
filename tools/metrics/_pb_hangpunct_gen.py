# -*- coding: utf-8 -*-
"""Which line-final punctuation may hang past the RIGHT MARGIN in Latin text?

`creative__009790431a821d2f` (EN Phase-1 FAIL, score 0.9680, 4 misplaced
paragraphs, no tables/images/footnotes/columns to confound) breaks one word
earlier than Word in three of its paragraphs. Read out of Word's own PDF, the
offending line is:

    content right edge      = 523.32       (A4, 72pt margins -> 451.32 column)
    ... Bankers Ghan|a|       523.38        <- fills the column exactly
    ... Bankers Ghana|.|      526.17        <- the PERIOD sits OUTSIDE it
    (trailing space)          528.56

So Word fits the text to the margin and lets the final period overhang; Oxi
counts the period in the wrap decision, breaks a word early and spends an extra
line. `w:overflowPunct` is ABSENT from that document's settings.xml, i.e. the
ECMA-376 default (true) is in force.

This probe isolates the rule and asks WHICH characters get the privilege. Each
arm is one paragraph of `x`-filler (Courier New, so every glyph has the same
advance and the boundary is exactly computable) plus one final punctuation mark,
with the filler length swept across the fit boundary. If the mark may hang, the
paragraph stays ONE line at the length where the filler alone exactly fills the
column; if not, it becomes two.

    python tools/metrics/_pb_hangpunct_gen.py         # build
    python tools/metrics/_pb_hangpunct_read.py word   # Word truth (COM -> PDF)
    python tools/metrics/_pb_hangpunct_read.py oxi    # Oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_hangpunct"
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""
SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:eastAsia="Courier New" w:cs="Courier New"/>
<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

# Courier New 10pt = 6.0pt per glyph. A4 with 72pt margins -> 451.32pt column
# -> 75 glyphs = 450.0 fit, 76 = 456.0 do not. Sweep N around that.
PUNCT = {
    "period": ".", "comma": ",", "semi": ";", "colon": ":",
    "bang": "!", "query": "?", "rparen": ")", "rbracket": "]",
    "rquote": "\u201d", "apos": "\u2019", "hyphen": "-",
    "letter": "z",          # control: a LETTER must never hang
    "none": "",             # control: filler only
}
NS = [74, 75, 76]
JC = {"left": "", "both": '<w:jc w:val="both"/>'}

ARMS = [("%s_%s_n%d" % (j, p, n), JC[j], PUNCT[p], n)
        for j in JC for p in PUNCT for n in NS]


def build(tag, jc, mark, n):
    text = "x" * n + mark
    body = ('<w:p><w:pPr>%s</w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (jc, text))
    body += '<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>'
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>\n" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
           ' w:header="708" w:footer="708" w:gutter="0"/>'
           "</w:sectPr></w:body></w:document>")
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    for a in ARMS:
        build(*a)
    print("built %d arms in %s  (column 451.32pt, Courier New 10 = 6.0pt/glyph)"
          % (len(ARMS), OUT))
