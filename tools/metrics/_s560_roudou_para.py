# -*- coding: utf-8 -*-
import sys,json,zipfile,re
from xml.etree import ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ns={'w':W[1:-1]}
z=zipfile.ZipFile('tools/golden-test/documents/docx/roudoujoken_001161383.docx')
root=ET.fromstring(z.read('word/document.xml'))
body=root.find('w:body',ns)
# all <w:p> in document order (incl table cells) = Word doc.Paragraphs order
paras=list(body.iter(W+'p'))
def ptext(p): return ''.join(e.text or '' for e in p.iter(W+'t'))
print('total <w:p>:',len(paras))
for i in (145,146,147,148,149,166,167,168):
    if i-1 < len(paras):
        print('  Word i=%-3d %r'%(i,ptext(paras[i-1])[:30]))
# Now find these in Oxi pagination
O=json.load(open(r'pipeline_data/pagination_oxi/roudoujoken.json',encoding='utf-8'))
def norm(s): return ''.join((s or '').split())[:16]
oxi_pos={}
for pg in O['pages']:
    for r in O['pages'][pg]:
        t=norm(r.get('text',''))
        if t and t not in oxi_pos: oxi_pos[t]=int(pg)
print('\nOxi page for each:')
for i in (146,147,148,166,167):
    t=norm(ptext(paras[i-1]))
    print('  Word i=%-3d oxi_page=%s  %r'%(i,oxi_pos.get(t,'?'),t))
