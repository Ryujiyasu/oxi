"""S492 follow-up 2 — realistic punct density: does Oxi's capacity-break over-pack
at the ~10-17% density of real Japanese prose (not the 50% of the canonical repro)?

Densities: d6 = 国x5 + 、 (1 punct / 6 = 17%); d10 = 国x9 + 、 (10%); d3 = 国x2+、 (33%).
Each at jc=both and jc=left. Measures Word L1 count + the 2nd-half punct advances.
"""
import zipfile, os
import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/breakflip_jc')
os.makedirs(OUT, exist_ok=True)
CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')


def docx(path, text, jc):
    body = ('<w:p><w:pPr><w:jc w:val="%s"/><w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>') % (jc, text)
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>') % body
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES); z.writestr('word/settings.xml', SETTINGS)


CASES = {
    'd3': '国国、' * 24,    # 33% punct
    'd6': '国国国国国、' * 12,   # 17% punct
    'd10': '国国国国国国国国国、' * 7,  # 10% punct
    # mixed brackets/comma like real prose
    'mix': '国国「国国」国、国国国。' * 6,
}
ALIGNS = ['both', 'left']
for k, t in CASES.items():
    for jc in ALIGNS:
        docx(os.path.join(OUT, 'rd_%s_%s.docx' % (k, jc)), t, jc)
print('generated', len(CASES) * len(ALIGNS), 'realistic-density variants')

WD_VPOS, WD_HPOS = 6, 5
word = w32.DispatchEx('Word.Application'); word.Visible = False
res = {}
try:
    for k in CASES:
        res[k] = {}
        for jc in ALIGNS:
            path = os.path.abspath(os.path.join(OUT, 'rd_%s_%s.docx' % (k, jc)))
            doc = word.Documents.Open(path, ReadOnly=True)
            rng = doc.Paragraphs(1).Range; text = rng.Text; start = rng.Start
            y0 = doc.Range(start, start).Information(WD_VPOS)
            xs = []
            for i in range(len(text)):
                ch = text[i]
                if ch in ('\r', '\n', '\x07'): continue
                if doc.Range(start + i, start + i).Information(WD_VPOS) > y0 + 2: break
                xs.append((ch, round(doc.Range(start + i, start + i).Information(WD_HPOS), 2)))
            advs = [round(xs[j + 1][1] - xs[j][1], 2) for j in range(len(xs) - 1)]
            ncomp = sum(1 for a in advs if a < 11.5)
            res[k][jc] = (len(xs), ncomp, len(advs), min(advs) if advs else 0)
            doc.Close(False)
finally:
    word.Quit()

print('\n=== Word L1 count (compressed mid-line advances <11.5 / total) ===')
print('%-6s %-22s %-22s' % ('case', 'jc=both', 'jc=left'))
for k in CASES:
    b = res[k]['both']; l = res[k]['left']
    print('%-6s L1=%2d comp=%2d/%2d min=%.2f   L1=%2d comp=%2d/%2d min=%.2f'
          % (k, b[0], b[1], b[2], b[3], l[0], l[1], l[2], l[3]))
