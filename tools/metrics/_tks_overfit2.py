# -*- coding: utf-8 -*-
import os, sys, json, tempfile, difflib, fitz, re
from collections import Counter
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
TMP=tempfile.gettempdir()
PDF=os.path.join(TMP,'tks_truth.pdf'); GLY=os.environ.get('GLY',os.path.join(TMP,'tks_gfit.json'))
WLO,WHI=int(os.environ.get('WLO',46)),int(os.environ.get('WHI',59))
isfoot=lambda t: bool(re.fullmatch(r'\s*\d{1,3}\s*', t)) or not t.strip()
def word_lines():
    out=[]
    for pno in range(WLO,WHI):
        rd=fitz.open(PDF)[pno-1].get_text('rawdict'); chs=[]
        for blk in rd['blocks']:
            if blk.get('type',0)!=0: continue
            for ln in blk.get('lines',[]):
                for sp in ln['spans']:
                    for c in sp['chars']:
                        b=c['bbox']
                        if 60<b[1]<790: chs.append((c['c'],b[0],b[2],(b[1]+b[3])/2))
        chs.sort(key=lambda t:(round(t[3]/3),t[1])); rows={}
        for c,x0,x1,y in chs: rows.setdefault(round(y/3),[]).append((c,x0,x1))
        for k in sorted(rows):
            row=sorted(rows[k],key=lambda t:t[1])
            if not isfoot(''.join(c for c,_,_ in row)): out.append(row)
    return out
def dec(s):
    try: return s.encode('latin1').decode('cp932')
    except: return s
def oxi_lines():
    raw=json.load(open(GLY,encoding='utf-8'))
    gl=raw if isinstance(raw,list) else (raw.get('glyphs') or [g for p in raw.get('pages',[]) for g in p.get('glyphs',[])])
    out=[];cur=[];lt=None
    for g in gl:
        c=dec(g['char']);top=g.get('top')
        if lt is None or abs(top-lt)<2: cur.append((c,g['x']))
        else: out.append(cur);cur=[(c,g['x'])]
        lt=top
    if cur: out.append(cur)
    return out
W=word_lines();O=oxi_lines()
def stream(lines):
    s=[];m=[]
    for li,ln in enumerate(lines):
        for j,it in enumerate(ln):
            if it[0].strip(): s.append(it[0]);m.append((li,j))
    return s,m
ws,wm=stream(W);os_,om=stream(O)
sm=difflib.SequenceMatcher(None,ws,os_,autojunk=False)
wl={}
for blk in sm.get_matching_blocks():
    for k in range(blk.size):
        wli,wj=wm[blk.a+k];oli,oj=om[blk.b+k]
        wc,wx0,wx1=W[wli][wj];oc,ox=O[oli][oj]
        wl.setdefault(wli,[]).append((wx0,wx1,oli,ox,wc))
def structural(txt):
    # numbering markers / heading brackets / table rows (tab-spaced)
    return txt.startswith('第') or txt.startswith('【') or txt.startswith('（') and '】' in txt
body=[];struct=[]
for wli in sorted(wl):
    m=wl[wli]
    if len(m)<8: continue
    olc=Counter(x[2] for x in m);dom=olc.most_common(1)[0][0]
    mm=[x for x in m if x[2]==dom]
    if len(mm)<8: continue
    # span = last-char-start - first-char-start (BOTH exclude last width)
    wspan=mm[-1][0]-mm[0][0]; ospan=mm[-1][3]-mm[0][3]
    drift=wspan-ospan; txt=''.join(x[4] for x in mm)
    (struct if structural(txt) else body).append((wli,len(mm),drift,txt))
def report(name,rows):
    rows=[r for r in rows]
    tot=sum(r[2] for r in rows); n=len(rows)
    print(f"\n=== {name}: {n} lines, total drift {tot:.1f}pt, mean {tot/max(1,n):.2f}pt/line, per-char {tot/max(1,sum(r[1] for r in rows)):.3f} ===")
    for wli,nn,drift,txt in sorted(rows,key=lambda r:-abs(r[2]))[:8]:
        print(f"   W{wli} n={nn} drift={drift:+.2f}  {txt[:32]}")
report("BODY lines (no marker/heading)",body)
report("STRUCTURAL lines (第N条/【】markers)",struct)
