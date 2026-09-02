# -*- coding: utf-8 -*-
"""Minimal repro: how do `w:cols` BANDS and CONTINUOUS section breaks compose
in a vertical (tbRl) section?

Derived from 015355870669f8d3 (Word COM per-character x/y walk) as:
  (a) `w:cols num=N` splits the content HEIGHT into N bands,
      band_h = (content_h - space*(N-1))/N;
  (b) a section owns the same COLUMN RANGE in every band, and the next section
      resumes after it -- a continuous break is a vertical cut across all bands
      (the horizontal continuous-multi-column rule rotated 90 degrees);
  (c) the columns of a section that ends in a continuous break are BALANCED
      over its bands, earlier bands taking the remainder;
  (d) the empty paragraph carrying the sectPr consumes NO column (S730's
      zero-height rule, in the other axis).

Each arm isolates one clause, and the arms are sized so that the balanced and
the fill-to-capacity predictions differ by SEVERAL columns -- a one-column
difference could be absorbed by a rounding disagreement about the character
advance, which would make the sweep unable to falsify anything.

Geometry (landscape A4, matching the real doc so the arithmetic carries over):
    page 841.9 x 595.3pt, margins T/B 85.05, L 85.05, R 99.25
    content height 425.2pt, text width 657.6pt
    every paragraph pinned to `exact` line 360tw => column pitch 18.0pt flat,
    so a page holds floor(657.6/18) = 36 columns and the column INDEX is
    readable straight off x:  col_k left edge = 742.65 - 18*k
    char advance down a column = 1em = 10.5pt
        1 band  : floor(425.20/10.5) = 40 chars per column
        2 bands : band_h 201.98 -> 19 chars
        3 bands : band_h 127.57 -> 12 chars

    python tools/metrics/_pb_vertcols_gen.py          # build
    python tools/metrics/_pb_vertcols_read.py         # Word truth (COM)
"""
import os
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_vertcols"
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
<w:rFonts w:ascii="\uff2d\uff33 \u660e\u671d" w:hAnsi="\uff2d\uff33 \u660e\u671d"
 w:eastAsia="\uff2d\uff33 \u660e\u671d" w:cs="\uff2d\uff33 \u660e\u671d"/>
<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="exact"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

# The section geometry every sectPr repeats. `cols` and `type` are per-arm.
SECT_TAIL = (
    '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    '<w:pgMar w:top="1701" w:right="1985" w:bottom="1701" w:left="1701"'
    ' w:header="851" w:footer="992" w:gutter="0"/>'
    '<w:textDirection w:val="tbRl"/>'
    '<w:docGrid w:type="lines" w:linePitch="360"/>'
)

# Fullwidth digits cycling 0-9 so the character INDEX of any glyph is readable
# off the glyph itself -- with 「あ」 repeated, a column boundary landing one
# character early or late is invisible.
DIGITS = "\uff10\uff11\uff12\uff13\uff14\uff15\uff16\uff17\uff18\uff19"


def text(n):
    return "".join(DIGITS[i % 10] for i in range(n))


def sect(cols_num, continuous, space=425):
    c = '<w:cols w:space="%d"/>' % space if cols_num == 1 else (
        '<w:cols w:num="%d" w:space="%d"/>' % (cols_num, space))
    t = '<w:type w:val="continuous"/>' if continuous else ""
    return "<w:sectPr>" + t + c + SECT_TAIL + "</w:sectPr>"


def para(s, sect_xml=None):
    ppr = "<w:pPr>%s</w:pPr>" % sect_xml if sect_xml else ""
    body = ('<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % s) if s else ""
    return "<w:p>%s%s</w:p>" % (ppr, body)


# ---------------------------------------------------------------- the arms --
# Each entry: (tag, [ (text, cols_of_the_section_this_para_ENDS, continuous) ],
#              final_section_cols, note)
# A paragraph with a non-None section entry carries that sectPr and so ENDS
# that section.
ARMS = {}

# A. bands: one section, 3 cols, 36 chars = exactly 3 columns of 12 in band 0.
#    Reads band_h straight off the y of the 13th and 25th character.
ARMS["a_bands3"] = ([(text(36), None)], 3, "36 chars, 3 bands -> 3 cols of 12 in band 0")

# B. cut: sec1 (1 col, 12 chars = 1 column), sec2 (3 cols, 12 chars = 1 column),
#    sec3 (1 col, 12 chars). If a continuous break is a column CUT, the three
#    land at columns 1, 2, 3. If sections just flowed, sec2's band grid would
#    restart at the right edge and they would collide.
ARMS["b_cut"] = ([(text(12), (1, False)), (text(12), (3, True)),
                  (text(12), (1, True))], 1, "1|3|1 cols, one column of text each")

# C. balance3: sec2 has 3 bands and 80 chars = ceil(80/12) = 7 columns.
#    BALANCED  -> 3+2+2, extent 3, so sec3 opens at column 5.
#    FILL      -> all 7 in band 0, extent 7, so sec3 opens at column 9.
ARMS["c_bal3"] = ([(text(12), (1, False)), (text(80), (3, True)),
                   (text(12), (1, True))], 1,
                  "3 bands, 7 columns: balanced -> sec3 at col 5, fill -> col 9")

# D. balance2: sec2 has 2 bands and 125 chars = ceil(125/19) = 7 columns.
#    BALANCED -> 4+3, extent 4, sec3 opens at column 6.
#    FILL     -> 7 in band 0, extent 7, sec3 opens at column 9.
ARMS["d_bal2"] = ([(text(12), (1, False)), (text(125), (2, True)),
                   (text(12), (1, True))], 1,
                  "2 bands, 7 columns: balanced -> sec3 at col 6, fill -> col 9")

# E. mark: does the EMPTY paragraph carrying the sectPr eat a column?
#    sec1 = 12 chars + an empty sectPr paragraph; sec2 = 12 chars.
#    consumes -> sec2 at column 3;  free (S730) -> column 2.
ARMS["e_mark"] = ([(text(12), None), ("", (1, False)), (text(12), (1, True))],
                  1, "empty sectPr para: eats a column -> sec2 at col 3, else col 2")

# F. mark_nonempty: the same, but the sectPr rides on a paragraph WITH text.
#    That one must always take its own column (sec2 at column 3).
ARMS["f_marktext"] = ([(text(12), None), (text(12), (1, False)),
                       (text(12), (1, True))], 1,
                      "sectPr on a NON-empty para -> always its own column")


def build(tag, paras, final_cols, note):
    xml = []
    for s, sec in paras:
        xml.append(para(s, sect(sec[0], sec[1]) if sec else None))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>\n" + "\n".join(xml) + sect(final_cols, False) + "</w:body></w:document>")
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
    for tag, (paras, fc, note) in ARMS.items():
        build(tag, paras, fc, note)
        print("  %-12s %s" % (tag, note))
    print("built %d arms in %s" % (len(ARMS), OUT))
