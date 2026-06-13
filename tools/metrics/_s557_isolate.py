# -*- coding: utf-8 -*-
"""S557 — verbatim-isolate d77a para9 (公共データ利用規約… the L6 divergent
para) into a minimal docx carrying d77a's settings.xml, styles.xml and
sectPr verbatim. Then mutations. Measures per-line char counts via COM.
Usage: python _s557_isolate.py [mutation]
  mutations: none | paren2toten (）→、) | move (shift punct) | font (gothic→mincho)
"""
import io
import os
import re
import sys
import zipfile

import win32com.client as w32

SRC = r'c:\tmp\_d77a_copy.docx'
OUT = os.path.abspath('tools/golden-test/repros/s557_isolate')
os.makedirs(OUT, exist_ok=True)

z = zipfile.ZipFile(SRC)
doc_xml = z.read('word/document.xml').decode('utf-8')
styles_xml = z.read('word/styles.xml').decode('utf-8')
settings_xml = z.read('word/settings.xml').decode('utf-8')

# locate the para containing the probe text (tag-stripped index map)
out = []
xmlpos = []
for m in re.finditer(r'<[^>]+>|[^<]+', doc_xml):
    s = m.group(0)
    if not s.startswith('<'):
        for k, ch in enumerate(s):
            out.append(ch)
            xmlpos.append(m.start() + k)
plain = ''.join(out)
i = plain.find(u'「公共データ利用規約（第1.0版）」の前身である')
xj = xmlpos[i]
ps = doc_xml.rfind('<w:p ', 0, xj)
pe = doc_xml.find('</w:p>', xj) + len('</w:p>')
para_xml = doc_xml[ps:pe]

# the LAST sectPr (body-level) — margins + grid for the main flow
sect_m = list(re.finditer(r'<w:sectPr[ >].*?</w:sectPr>', doc_xml, re.S))[-1]
sect_xml = sect_m.group(0)
sect_xml = re.sub(r'<w:headerReference[^>]*/>', '', sect_xml)
sect_xml = re.sub(r'<w:footerReference[^>]*/>', '', sect_xml)

# sanitize: drop cross-paragraph / relationship-bearing fragments
for pat in (r'<w:bookmarkStart[^>]*/>', r'<w:bookmarkEnd[^>]*/>',
            r'<w:proofErr[^>]*/>', r'<w:commentRange\w*[^>]*/>',
            r'<w:commentReference[^>]*/>'):
    para_xml = re.sub(pat, '', para_xml)
settings_xml = re.sub(r'<w:attachedTemplate[^>]*/>', '', settings_xml)
settings_xml = re.sub(r'<w:rsids>.*?</w:rsids>', '', settings_xml, flags=re.S)
# sectPr may carry header/footer references -> strip them
sect_clean_holder = []

mutation = sys.argv[1] if len(sys.argv) > 1 else 'none'
if mutation == 'paren2toten':
    para_xml = para_xml.replace(u'認める）形', u'認める、形')
elif mutation == 'nolatin':
    para_xml = para_xml.replace(u'1.0', u'一〇')

# Clone the WHOLE d77a package (all parts/relationships intact) and replace
# only document.xml's body with the single para + the original body sectPr
# (header/footer refs kept — the parts exist in the clone).
sect_m2 = list(re.finditer(r'<w:sectPr[ >].*?</w:sectPr>', doc_xml, re.S))[-1]
sect_orig = sect_m2.group(0)
root_end = doc_xml.find('<w:body>') + len('<w:body>')
doc_head = doc_xml[:root_end]
new_doc = doc_head + para_xml + sect_orig + '</w:body></w:document>'

docx = os.path.join(OUT, 's557_%s.docx' % mutation)
with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as zz:
    for item in z.infolist():
        if item.filename == 'word/document.xml':
            zz.writestr('word/document.xml', new_doc)
        else:
            zz.writestr(item, z.read(item.filename))

word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    try:
        pr = wdoc.Paragraphs(1).Range
        start = pr.Start
        y0 = None
        cnt = 0
        lines = []
        for i in range(min(pr.End - pr.Start, 420)):
            ch = wdoc.Range(start + i, start + i + 1).Text or ''
            if ch in ('\r', '\x07', '\n'):
                continue
            y = wdoc.Range(start + i, start + i).Information(6)
            if y0 is None:
                y0 = y
            if abs(y - y0) > 0.5:
                lines.append(cnt)
                cnt = 0
                y0 = y
            cnt += 1
        lines.append(cnt)
        sys.stdout.reconfigure(encoding='utf-8')
        print('%s lines: %s' % (mutation, lines))
    finally:
        wdoc.Close(False)
finally:
    word.Quit()
