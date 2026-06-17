# -*- coding: utf-8 -*-
"""Find the EXACT char where Oxi gets 1 page ahead of Word (the -1 onset),
with para/line context. Reuses _tks_oidashi-style char-stream alignment."""
import os, sys, json, tempfile, difflib, fitz
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
TMP=tempfile.gettempdir()
PDF=os.path.join(TMP,'tks_truth.pdf')
DUMP=os.environ.get('TKS_DUMP', os.path.join(TMP,'tks_s586.json'))
WLO,WHI=int(os.environ.get('WLO',44)),int(os.environ.get('WHI',62))

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
                        if y0<60 or y0>800: continue
                        chars.append((ch['c'],b[0],b[2],(b[1]+b[3])/2))
        chars.sort(key=lambda t:(round(t[3]/3.0),t[1]))
        rows={}
        for c,x0,x1,yc in chars: rows.setdefault(round(yc/3.0),[]).append((c,x0,x1,yc))
        for k in sorted(rows):
            row=sorted(rows[k],key=lambda t:t[1])
            out.append({'page':pno,'chars':[c for c,_,_,_ in row]})
    return out

def oxi_lines():
    d=json.load(open(DUMP,encoding='utf-8')); out=[]
    for pg in d['pages']:
        if not (WLO<=pg['page']<WHI+2): continue
        rows={}
        for el in pg.get('elements',[]):
            if el.get('type')!='text' or not el.get('text'): continue
            rows.setdefault(round(el['y'],0),[]).append(el)
        for k in sorted(rows):
            row=sorted(rows[k],key=lambda e:e['x'])
            txt=''.join(e.get('text','') for e in row)
            out.append({'page':pg['page'],'text':txt})
    return out

W=word_lines(); O=oxi_lines()
def stream(lines,gc):
    s=[]; idx=[]
    for li,ln in enumerate(lines):
        for c in gc(ln):
            if c.strip(): s.append(c); idx.append(li)
    return s,idx
ws,wi=stream(W,lambda l:l['chars'])
os_,oi=stream(O,lambda l:list(l['text']))
sm=difflib.SequenceMatcher(None,ws,os_,autojunk=False)
# walk matched chars, report each time the running delta changes
prev=None
for blk in sm.get_matching_blocks():
    for k in range(blk.size):
        a=blk.a+k; b=blk.b+k
        wp=W[wi[a]]['page']; op=O[oi[b]]['page']
        d=op-wp
        if d!=prev:
            ctx=''.join(ws[max(0,a-12):a+1])
            print(f"  Δ{d:+d}  Wp{wp}->Op{op}  ...{ctx}")
            prev=d
