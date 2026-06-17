# -*- coding: utf-8 -*-
import os, sys, json, tempfile, difflib, fitz, re
from collections import Counter
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
TMP=tempfile.gettempdir()
PDF=os.path.join(TMP,'tks_truth.pdf')
DUMP=os.environ.get('TKS_DUMP', os.path.join(TMP,'tks_s586.json'))
WLO,WHI=int(os.environ.get('WLO',44)),int(os.environ.get('WHI',64))
isfoot=lambda t: bool(re.fullmatch(r'\s*\d{1,3}\s*', t))   # footer page-number line

def word_lines():
    doc=fitz.open(PDF); out=[]
    for pno in range(WLO,WHI):
        pg=doc[pno-1]; rd=pg.get_text('rawdict'); chars=[]
        for blk in rd['blocks']:
            if blk.get('type',0)!=0: continue
            for ln in blk.get('lines',[]):
                for sp in ln['spans']:
                    for ch in sp['chars']:
                        b=ch['bbox']; y0=b[1]
                        if y0<60 or y0>788: continue
                        chars.append((ch['c'],b[0],b[2],(b[1]+b[3])/2))
        chars.sort(key=lambda t:(round(t[3]/3.0),t[1]))
        rows={}
        for c,x0,x1,yc in chars: rows.setdefault(round(yc/3.0),[]).append((c,x0,x1,yc))
        for k in sorted(rows):
            row=sorted(rows[k],key=lambda t:t[1]); txt=''.join(c for c,_,_,_ in row)
            if isfoot(txt): continue
            out.append({'page':pno,'chars':list(txt)})
    return out
def oxi_lines():
    d=json.load(open(DUMP,encoding='utf-8')); out=[]
    for pg in d['pages']:
        rows={}
        for el in pg.get('elements',[]):
            if el.get('type')!='text' or not el.get('text'): continue
            rows.setdefault(round(el['y'],0),[]).append(el)
        for k in sorted(rows):
            row=sorted(rows[k],key=lambda e:e['x']); txt=''.join(e.get('text','') for e in row)
            if isfoot(txt): continue
            out.append({'page':pg['page'],'chars':list(txt)})
    return out
W=word_lines(); O=oxi_lines()
def stream(lines):
    s=[]; idx=[]
    for li,ln in enumerate(lines):
        for c in ln['chars']:
            if c.strip(): s.append(c); idx.append(li)
    return s,idx
ws,wi=stream(W); os_,oi=stream(O)
sm=difflib.SequenceMatcher(None,ws,os_,autojunk=False)
perpage={}
for blk in sm.get_matching_blocks():
    for k in range(blk.size):
        wp=W[wi[blk.a+k]]['page']; op=O[oi[blk.b+k]]['page']
        perpage.setdefault(wp,Counter())[op-wp]+=1
for wp in sorted(perpage):
    if not (WLO<=wp<WHI): continue
    c=perpage[wp]; dom=c.most_common(1)[0][0]
    print(f"  Wp{wp}: dom Δ={dom:+d}  {dict(sorted(c.items()))}")
