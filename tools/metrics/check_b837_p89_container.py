# -*- coding: utf-8 -*-
"""S492c — what structure contains the 本府は paragraph in b837's document.xml?
If it's inside w:txbxContent / mc:AlternateContent / w:sdt, Oxi's different handling
is structural (explains the missing content). If it's a plain w:body paragraph that
Oxi drops, that's a real parse bug worth flagging.
"""
import zipfile, glob, re

DOCX = glob.glob('tools/golden-test/documents/docx/b837*.docx')[0]
doc = zipfile.ZipFile(DOCX).read('word/document.xml').decode('utf-8')

# locate 本府 (may be split across runs); find by joined-text para then locate its span
paras = list(re.finditer(r'<w:p\b[^>]*>.*?</w:p>', doc, re.S))
target = None
for m in paras:
    body = m.group(0)
    txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', body, re.S))
    if '本府' in txt:
        target = m
        print("found 本府 para at document.xml char %d, text[:30]=%r" % (m.start(), txt[:30]))
        break

if target is None:
    print("本府 para not found in document.xml w:p list")
else:
    # walk the prefix before the para; count open/unclosed structural tags
    prefix = doc[:target.start()]
    for tag in ['w:tbl', 'w:tc', 'w:txbxContent', 'mc:AlternateContent', 'mc:Choice',
                'mc:Fallback', 'w:sdtContent', 'v:textbox', 'w:ftr', 'w:hdr']:
        opens = len(re.findall(r'<%s[\s>]' % re.escape(tag), prefix))
        closes = len(re.findall(r'</%s>' % re.escape(tag), prefix))
        if opens - closes != 0:
            print("  INSIDE open <%s> (open=%d close=%d, net=%d)" % (tag, opens, closes, opens - closes))
    print("  (any 'INSIDE' line above = the para is nested in that structure)")
