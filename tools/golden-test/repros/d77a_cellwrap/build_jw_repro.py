# -*- coding: utf-8 -*-
"""S434/S435: minimal repro generator for the d77a MS PGothic cell-wrap bug.

Extracts the d77a "カ" list paragraph (item J, Word i=138) verbatim and emits
single-1x1-table docx variants at a sweep of cell widths. At cell width ~8600tw
(margin-free), Word wraps J to 2 lines but Oxi keeps it on 1 — the boundary that
reproduces the bug. Root cause (S435): Oxi's com_twips_widths MS PGothic 10.5pt
advance widths are ~0.2pt/char too narrow vs Word, so J's 42-char body sums to
396.5pt in Oxi vs Word's ~405pt, fitting one extra char per line.

Run: python tools/golden-test/repros/d77a_cellwrap/build_jw_repro.py
"""
import re, os, zipfile
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
DOCX = os.path.join(REPO, "tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
OUT = os.path.dirname(__file__)

x = zipfile.ZipFile(DOCX).read('word/document.xml').decode('utf-8')
marker = body = None
for m in re.finditer(r'<w:p[ >].*?</w:p>', x, re.S):
    t = re.sub(r'<[^>]+>', '', ''.join(re.findall(r'<w:t[ >][^<]*</w:t>', m.group(0))))
    if 'ウェブサイト' in t and '全体' in t and 30 < len(t) < 55:
        marker, body = t[0], t[1:]; break
assert marker, "J paragraph not found"

CT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
ST = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ Ｐゴシック" w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="ＭＳ Ｐゴシック"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'

def esc(s): return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def docx(w, mar):
    cm = '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>' if mar == 0 else ''
    para = '<w:p><w:pPr><w:ind w:leftChars="202" w:left="564" w:hanging="140"/></w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (esc(marker), esc(body))
    tbl = '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>%s<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>' % (w, cm, w, w, para)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + tbl + '<w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>'

def write(name, w, mar):
    with zipfile.ZipFile(os.path.join(OUT, name), 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DR); z.writestr('word/styles.xml', ST)
        z.writestr('word/document.xml', docx(w, mar))

for w in (9072, 8952, 8800, 8600, 8400):
    write('JW_%d.docx' % w, w, None)
write('JW_8816_nomar.docx', 8816, 0)
write('JW_8600_nomar.docx', 8600, 0)  # Word=2 lines, Oxi=1 line — the bug
print("built. Key repro: JW_8600_nomar.docx (Word 2 lines, Oxi 1 line)")
