# -*- coding: utf-8 -*-
"""How does Word quantise a line's position, and does the page top share the grid?

Reading `creative__0158c02ae543d567`'s truth PDF showed the line pitch takes only
two values -- 15.840 and 15.960, 0.12pt = 1/600 inch apart, mixed 110:33 so the
mean is the exact 15.869 -- and that a 13-line paragraph's cumulative deviation
from the exact position stays inside +-0.06pt without drifting. That is
cumulative rounding onto a 1/600-inch grid. Two things it does not settle:

  ROUNDING  cumulative (deviation bounded by half a unit) or per-pitch
            (deviation is a random walk that grows with N)? One long paragraph
            of N lines answers it: fit y_n - y_0 against n * P_exact.
  ORIGIN    is the FIRST baseline on the same grid? Sweeping the top margin in
            1 twip (0.05pt) steps answers it directly: a quantised origin moves
            in 0.12pt jumps, an exact one moves 0.05 at a time.

The pitch is set by the font size and the line multiple, so several arms give
the grid different fractional parts to resolve.

    python tools/metrics/_pb_linequant_gen.py [--sweep lo hi step]
    python tools/metrics/_pb_linequant_read.py word|oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = r"C:\tmp\pb_linequant"
W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '</Relationships>')
# (label, font, half-point size, w:line value) -- pitches with different
# fractional parts of the 0.12pt unit.
FACES = [
    ("ar12x115", "Arial", 24, 276),          # 15.869 -> 132.24 units
    ("ar12x100", "Arial", 24, 240),          # 13.799 -> 114.99 units
    ("cal11x100", "Calibri", 22, 240),       # 13.428 -> 111.90 units
    ("tnr12x150", "Times New Roman", 24, 360),
]
NLINES = int(os.environ.get("LQ_LINES", "40"))
WORD_ = "quantisation "


def styles(font, sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles ' + W_NS + '>'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/><w:sz w:val="%d"/>'
            % (font, font, font, sz) +
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/></w:style></w:styles>')


def build(tag, font, sz, line, top_tw):
    text = ("L " + WORD_ * 14) * NLINES        # one paragraph, many lines
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document ' + W_NS + '><w:body>'
           '<w:p><w:pPr><w:spacing w:after="0" w:line="%d" w:lineRule="auto"/>'
           '</w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (line, text) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="%d" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/>' % top_tw +
           '</w:sectPr></w:body></w:document>')
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles(font, sz))
        z.writestr("word/document.xml", doc)
    return path


def parse_sweep(argv):
    if "--sweep" in argv:
        i = argv.index("--sweep")
        return list(range(int(argv[i + 1]), int(argv[i + 2]) + 1, int(argv[i + 3])))
    return list(range(1440, 1453))        # 13 x 1tw = 0.6pt of top margin


def arms(sweep):
    return [("%s_t%d" % (lab, t), font, sz, line, t)
            for lab, font, sz, line in FACES for t in sweep]


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    sw = parse_sweep(sys.argv)
    a = arms(sw)
    for t in a:
        build(*t)
    print("built %d arms (top margin %d..%d tw) in %s" % (len(a), sw[0], sw[-1], OUT))
