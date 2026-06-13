# -*- coding: utf-8 -*-
"""S561 real-doc differential: roudoujoken page-2 paragraph GAPS, Word vs Oxi,
matched by TEXT. The reference-point offset (Word line-box-top vs Oxi glyph-top)
cancels in consecutive gaps -> a paragraph where Word_gap >> Oxi_gap is where
Oxi under-counts height (the ~15.7pt that lets page 3 start higher)."""
import json,sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
def dec(s):
    try: return s.encode('latin1').decode('utf-8')
    except: return s
def norm(s): return ''.join((s or '').split())
W=json.load(open(r'pipeline_data/pagination_word/roudoujoken.json',encoding='utf-8'))
d=json.load(open(r'C:/Users/ryuji/AppData/Local/Temp/roudou_layout.json',encoding='utf-8'))
# Oxi: para_idx -> (page, top_y, joined_text)
opar=defaultdict(lambda:[None,1e9,[]])
for pgno,p in enumerate(d['pages']):
    for e in p['elements']:
        if e['type']!='text' or not e['text'].strip(): continue
        pi=e.get('para_idx')
        if pi is None: continue
        r=opar[pi]
        if r[0] is None: r[0]=pgno+1
        if e['y']<r[1]: r[1]=e['y']
        r[2].append((e['y'],e['x'],e['text']))
# build oxi text per para (top line only, sorted by x)
oxi=[]
for pi in sorted(opar):
    pg,ty,elems=opar[pi]
    topline=sorted([e for e in elems if abs(e[0]-ty)<2],key=lambda e:e[1])
    txt=''.join(e[2] for e in topline)
    oxi.append((pi,pg,ty,norm(txt)[:12]))
oxi_by_t={}
for pi,pg,ty,t in oxi:
    if t and t not in oxi_by_t: oxi_by_t[t]=(pg,ty)
# Word page2+3 paras with decoded text
print('Word vs Oxi page-2/3 para gaps (matched by text):')
print('Wi  Wpg Wy     Wgap  | Opg Oy     Ogap  | dGap   text')
prevW=prevO=None
for p in W['paragraphs']:
    if p['page'] not in (2,3): continue
    wt=norm(dec(p.get('text','')))[:12]
    o=oxi_by_t.get(wt)
    wy=p['y']; wgap=(wy-prevW) if (prevW is not None) else 0
    if o:
        opg,oy=o; ogap=(oy-prevO) if (prevO is not None and prevO[0]==opg) else 0
        dg=wgap-ogap
        flag=' <<<' if abs(dg)>=4 and prevW is not None else ''
        print('%-3d p%d %6.1f %5.1f | p%d %6.1f %5.1f | %+5.1f %s%s'%(p['i'],p['page'],wy,wgap,opg,oy,ogap,dg,(dec(p['text'])[:16] or '(empty)'),flag))
        prevO=(opg,oy)
    else:
        print('%-3d p%d %6.1f %5.1f | (no oxi match)        | %s'%(p['i'],p['page'],wy,wgap,dec(p['text'])[:16] or '(empty)'))
    prevW=wy
