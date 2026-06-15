# -*- coding: utf-8 -*-
"""Build a minimal type=lines linePitch=360 docx to isolate Word's grid-snap
behaviour for MULTIPLE line-spacing (lineRule="auto", line != 240) BODY paras.

Each test paragraph holds one short line of MS Mincho 10.5pt text and a given
line=N auto spacing, bracketed by 10.5pt single-spaced anchors (each anchor = 1
grid cell = 18pt). The test para's advance = Y(anchor_after) - Y(test).

Discriminator:
  line=204 (0.85x): natural ~11.5 -> 18 if Word floors to >=1 cell
  line=300 (1.25x): natural ~16.9 -> 18 (1 cell) either way
  line=360 (1.50x): natural ~20.2 -> 20.x if "min 1 cell" / 36 if "ceil to cells"
  line=420 (1.75x): natural ~23.6 -> 23.x if "min 1 cell" / 36 if "ceil to cells"
  line=480 (2.00x): natural ~27.0 -> 27.x if "min 1 cell" / 36 if "ceil to cells"
"""
import os, zipfile

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

def anchor(n):
    return ('<w:p><w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/>'
            '<w:sz w:val="21"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'
            f'<w:t>ANCHOR{n}</w:t></w:r></w:p>')

def test_para(line, text, sz=21):
    return ('<w:p><w:pPr>'
            f'<w:spacing w:line="{line}" w:lineRule="auto"/>'
            '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>')

# (label, line, sz). 240 = single control. 21=10.5pt, 28=14pt.
variants = [
    ("L204", 204, 21),
    ("L240", 240, 21),
    ("L300", 300, 21),
    ("L360", 360, 21),
    ("L420", 420, 21),
    ("L480", 480, 21),
    ("L240_14", 240, 28),
    ("L300_14", 300, 28),
    ("L360_14", 360, 28),
    ("L480_14", 480, 28),
]
# also an EMPTY line=360 test (mirrors 3a4f/model body para)
def test_empty(line):
    return ('<w:p><w:pPr>'
            f'<w:spacing w:line="{line}" w:lineRule="auto"/>'
            '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/>'
            '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr></w:p>')

body = []
for i, (label, line, sz) in enumerate(variants):
    body.append(anchor(f"{i}A"))
    body.append(test_para(line, f"行{label}", sz))
    body.append(anchor(f"{i}B"))
# empty line=360 test
body.append(anchor("E0A"))
body.append(test_empty(360))
body.append(anchor("E0B"))

sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')

doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<w:document {NS}><w:body>' + ''.join(body) + sect + '</w:body></w:document>')

ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

out = os.path.join(os.path.dirname(__file__), "..", "golden-test", "repros",
                   "mult_grid", "mult_grid.docx")
os.makedirs(os.path.dirname(out), exist_ok=True)
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", ct)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", doc)
print("wrote", os.path.abspath(out))
print("variants:", [v[0] for v in variants], "+ empty L360")
