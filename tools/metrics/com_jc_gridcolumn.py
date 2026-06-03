# -*- coding: utf-8 -*-
"""S492x — decisive: is Word's {9.75 base, periodic 10.5} cell char pattern from GRID-COLUMN
quantization (would appear on jc=left too) or from jc=both JUSTIFY (discrete slack)? Generate
2 controlled single-cell docs (10.5pt, docGrid linesAndChars charSpace=-2714, wrapping CJK):
one jc=both, one jc=left. COM-measure char advance pattern on a NON-last (full) line of each.
If jc=left also shows {9.75/10.5} stepping -> grid quantization (fix=column snap). If jc=left
is uniform ~9.84 and only jc=both steps -> discrete justify (fix=quantize justify slack).
cp932-safe (UTF-8 file). Writes ASCII results."""
import os, zipfile, json
import win32com.client as win32

OUT = r'c:\tmp\jcgrid'
os.makedirs(OUT, exist_ok=True)
wdHorizPos = 5
wdVertPos = 6

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOCRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''
SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:adjustLineHeightInTable/></w:settings>'''
TEXT = '東京都の電子計算機処理等に係る個人情報の保護に関する管理規程について定めるものとする'


def make(jc):
    rpr = '<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
    para = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="%s"/>%s</w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (jc, rpr, rpr, TEXT * 3))
    cell = '<w:tc><w:tcPr><w:tcW w:w="4200" w:type="dxa"/></w:tcPr>%s</w:tc>' % para
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="4200" w:type="dxa"/><w:tblBorders>'
           '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="4200"/></w:tblGrid><w:tr>%s</w:tr></w:tbl>' % cell)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>%s<w:p/>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/></w:sectPr></w:body></w:document>' % tbl)


paths = {}
for jc in ['both', 'left']:
    p = os.path.join(OUT, 'jc_%s.docx' % jc)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOCRELS)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', make(jc))
    paths[jc] = p

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
res = {}
try:
    for jc, p in paths.items():
        doc = word.Documents.Open(p, ReadOnly=True)
        try:
            rng = doc.Tables(1).Cell(1, 1).Range
            start, end = rng.Start, rng.End - 1
            pts = []
            for j in range(start, end):
                cr = doc.Range(j, j)
                try:
                    pts.append((round(float(cr.Information(wdHorizPos)), 2), round(float(cr.Information(wdVertPos)), 1)))
                except Exception:
                    pass
            # group by line (y), take the FIRST full line's advances
            from collections import defaultdict, OrderedDict
            byline = OrderedDict()
            for x, y in pts:
                byline.setdefault(y, []).append(x)
            res[jc] = []
            for y, xs in byline.items():
                xs = sorted(xs)
                advs = [round(xs[k] - xs[k - 1], 2) for k in range(1, len(xs))]
                res[jc].append((y, len(xs), advs))
        finally:
            doc.Close(False)
finally:
    word.Quit()

print("Word cell char advance pattern (10.5pt, linesAndChars charSpace=-2714, adjustLineHeightInTable):")
for jc in ['both', 'left']:
    print("\n=== jc=%s ===" % jc)
    for y, n, advs in res[jc][:4]:
        from collections import Counter
        print("  line y=%.1f n=%d advs(uniq counts)=%s" % (y, n, dict(Counter(advs))))
print("\nVERDICT: if jc=left ALSO steps {9.75/10.5} -> GRID quantization (Oxi uniform 9.84 is the bug, fix=column snap).")
print("  if jc=left is uniform and only jc=both steps -> DISCRETE JUSTIFY (fix=quantize justify slack).")
with open(r'c:\tmp\jcgrid_result.json', 'w', encoding='utf-8') as f:
    json.dump(res, f, ensure_ascii=False, indent=1)
