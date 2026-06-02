"""S492 (2026-06-03) — jc=left disentanglement of break-decision vs justify-render.

The R35 multi-session refactor (user option A). EVERY prior breakflip/d77a/b837
repro used jc=both, which CONFOUNDS two variables:
  (1) the BREAK decision width (how many chars Word fits on line 1), and
  (2) the JUSTIFY render width (how Word spreads/compresses to fill the margin).

This script regenerates the 6 canonical breakflip cases (国 + punct alternating,
MS Mincho 12pt, docGrid type=lines, avail 453.5pt, compressPunctuation, compat 15)
across 4 alignments {both,left,right,center} and measures, per variant:
  - L1 char count  (= the break decision, alignment-independent IF break is
                     done at fixed widths regardless of jc)
  - L1 per-char advances of the first ~10 chars (= the render width; for jc!=both
                     there is NO post-break spreading, so this is the natural or
                     break-time-compressed width directly).

DECISIVE outcome:
  - If jc=left STILL fits 38 (国、) and renders 、 < 12.0 -> punct compression is a
    real BREAK-TIME mechanism, independent of justify. The break model must
    compress punct for ALL alignments.
  - If jc=left drops to 37 and renders 、 = 12.0 (natural) -> the "+1" under jc=both
    is a JUSTIFY-coupled effect. Non-justified paragraphs break at NATURAL widths;
    only jc=both/distribute get the punct-compression break. This massively
    simplifies the Phase-1-sensitive break model (scope compression to justified).

Prior jc=both baseline (repros/breakflip, session471): kanji=37, comma=38,
period=38, close_kak=38, close_paren=38, open_kak=39.
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
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)


CASES = {
    'kanji': '国' * 60,            # 国 baseline (no punct)
    'comma': '国、' * 30,      # 国、
    'period': '国。' * 30,     # 国。
    'close_kak': '国」' * 30,  # 国」
    'open_kak': '国「' * 30,   # 国「  (opening; line-end-prohibited)
    'close_paren': '国）' * 30,  # 国）
}
ALIGNS = ['both', 'left', 'right', 'center']

for k, t in CASES.items():
    for jc in ALIGNS:
        docx(os.path.join(OUT, 'bf_%s_%s.docx' % (k, jc)), t, jc)
print('generated', len(CASES) * len(ALIGNS), 'docx variants in', OUT)

# ---- measure ----
WD_VPOS = 6  # wdVerticalPositionRelativeToPage
WD_HPOS = 5  # wdHorizontalPositionRelativeToPage
word = w32.DispatchEx('Word.Application')
word.Visible = False
results = {}
try:
    for k in CASES:
        results[k] = {}
        for jc in ALIGNS:
            path = os.path.abspath(os.path.join(OUT, 'bf_%s_%s.docx' % (k, jc)))
            doc = word.Documents.Open(path, ReadOnly=True)
            rng = doc.Paragraphs(1).Range
            text = rng.Text
            start = rng.Start
            # collapsed-range Information (R30): query at doc.Range(p,p)
            y0 = doc.Range(start, start).Information(WD_VPOS)
            xs = []  # (char, x) for chars on L1, in order
            n = 0
            for i in range(len(text)):
                ch = text[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                x = doc.Range(start + i, start + i).Information(WD_HPOS)
                if doc.Range(start + i, start + i).Information(WD_VPOS) > y0 + 2:
                    break
                xs.append((ch, x))
                n += 1
            # per-char advances of first ~10 chars (x[i+1]-x[i]); last one unknown
            adv = []
            for j in range(min(11, len(xs)) - 1):
                adv.append((xs[j][0], round(xs[j + 1][1] - xs[j][1], 2)))
            results[k][jc] = {'L1': n, 'adv': adv}
            doc.Close(False)
finally:
    word.Quit()

# ---- report ----
print('\n=== L1 char count by alignment (avail 453.5pt; pure-kanji natural=37) ===')
print('%-12s %6s %6s %6s %6s' % ('case', 'both', 'left', 'right', 'center'))
for k in CASES:
    r = results[k]
    print('%-12s %6d %6d %6d %6d' % (k, r['both']['L1'], r['left']['L1'], r['right']['L1'], r['center']['L1']))

print('\n=== L1 punct advance (the 2nd char; natural fullwidth = 12.0) ===')
print('%-12s %14s %14s %14s %14s' % ('case', 'both', 'left', 'right', 'center'))
for k in CASES:
    if k == 'kanji':
        continue
    r = results[k]
    def punct_adv(jc):
        a = r[jc]['adv']
        # 2nd char is the punct (index 1)
        return a[1][1] if len(a) > 1 else None
    print('%-12s %14s %14s %14s %14s' % (k, punct_adv('both'), punct_adv('left'), punct_adv('right'), punct_adv('center')))

print('\n=== full L1 advance traces (first 10 chars) ===')
import json
print(json.dumps(results, ensure_ascii=False, indent=1))
