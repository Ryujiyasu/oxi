# -*- coding: utf-8 -*-
"""S558 — opening-bracket compression spec. Is the （ trim a FIXED kern
(always -1.5 after content) or JUSTIFICATION compression (only when the line
is overset, capped ~1.5pt)? Measure （ advance in:
  A. （ after kanji, line NOT overset (short, lots of room)
  B. （ after kanji, line overset (needs compression to fit)
  C. （ after 」 (closing bracket)
  D. （ after ） (closing bracket)
  E. （ after 。
  F. 「 after kanji (another opener)
all at c15 jc=both fs12. Advance = Information(5) delta. If A=12.0 and
B<12.0 -> justification compression (variable). If both <12.0 -> fixed kern.
"""
import os
import sys
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s558_opener')
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
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
          '<w:kern w:val="2"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def build(docx, text, right_mar):
    body = ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="%d" w:bottom="1134" w:left="1304"/></w:sectPr>' % right_mar)
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (body, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)


# each case: (name, text). The text is engineered so the FIRST line contains
# the probe opener. For "overset" we make a long line that must compress; for
# "loose" a short line with room. Probe = the opener char; we report its
# advance and the char before it.
TAIL = u'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも'
CASES = {
    # loose: short line, big right margin -> line not full -> no compression
    'A_kanji_loose':  u'国国国（国国国',
    'C_close_loose':  u'国国「国」（国国',
    'E_period_loose': u'国国国。（国国',
    'F_kaku_loose':   u'国国国「国国国',
}
# overset cases: 38 chars with the opener early, tight margin
CASES_OVER = {
    'A_kanji_over':  u'国（' + u'国' * 36 + TAIL,   # （ after 国
    'C_close_over':  u'国」（' + u'国' * 35 + TAIL,  # （ after 」
    'D_paren_over':  u'国）（' + u'国' * 35 + TAIL,  # （ after ）
    'E_period_over': u'国。（' + u'国' * 35 + TAIL,  # （ after 。
    'F_kaku_over':   u'国「' + u'国' * 36 + TAIL,    # 「 after 国
}


def measure(word, docx, probe_pos):
    """advance of the char at index probe_pos (0-based within para)."""
    wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    try:
        pr = wdoc.Paragraphs(1).Range
        s = pr.Start
        x0 = wdoc.Range(s + probe_pos, s + probe_pos).Information(5)
        x1 = wdoc.Range(s + probe_pos + 1, s + probe_pos + 1).Information(5)
        ch = wdoc.Range(s + probe_pos, s + probe_pos + 1).Text
        return ch, (x1 - x0)
    finally:
        wdoc.Close(False)


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    sys.stdout.reconfigure(encoding='utf-8')
    print('=== LOOSE (right margin 6000tw, line not full) ===')
    for name, text in CASES.items():
        docx = os.path.join(OUT, '%s.docx' % name)
        build(docx, text, 6000)
        # opener is at index 3 for A/E/F (国国国X), index 4 for C (国国「国」-> ( at 5)
        pos = text.index(u'（') if u'（' in text else text.index(u'「')
        ch, adv = measure(word, docx, pos)
        print('  %-16s probe[%d]=%s adv=%.2f' % (name, pos, ch, adv))
    print('=== OVERSET (right margin tight, line must compress) ===')
    for name, text in CASES_OVER.items():
        docx = os.path.join(OUT, '%s.docx' % name)
        build(docx, text, 1500)
        pos = text.index(u'（') if u'（' in text else text.index(u'「')
        ch, adv = measure(word, docx, pos)
        # also the char before
        chp, advp = measure(word, docx, pos - 1)
        print('  %-16s prev[%d]=%s adv=%.2f | probe[%d]=%s adv=%.2f'
              % (name, pos - 1, chp, advp, pos, ch, adv))
finally:
    word.Quit()
