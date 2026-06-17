# -*- coding: utf-8 -*-
"""Characterize tokyoshugyo #2: per-line cumulative WIDTH drift Oxi vs Word on
the 賃金 chapter. For each Word visual line, gather the matched Oxi chars
(char-stream difflib align), compute Word span vs Oxi span of the SAME chars.
drift = Word_span - Oxi_span (>0 = Oxi narrower = can over-fit). Aggregates the
drift and attributes it to char classes."""
import os, sys, json, tempfile, difflib, fitz, re
from collections import Counter
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
TMP=tempfile.gettempdir()
PDF=os.path.join(TMP,'tks_truth.pdf')
GLY=os.environ.get('GLY', os.path.join(TMP,'tks_gfit.json'))
WLO,WHI=int(os.environ.get('WLO',46)),int(os.environ.get('WHI',59))
isfoot=lambda t: bool(re.fullmatch(r'\s*\d{1,3}\s*', t)) or not t.strip()

# Word: per-char (c,x0,x1) clustered into lines
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
        chs.sort(key=lambda t:(round(t[3]/3),t[1]))
        rows={}
        for c,x0,x1,y in chs: rows.setdefault(round(y/3),[]).append((c,x0,x1))
        for k in sorted(rows):
            row=sorted(rows[k],key=lambda t:t[1]); t=''.join(c for c,_,_ in row)
            if not isfoot(t): out.append(row)
    return out

def dec(s):
    try: return s.encode('latin1').decode('cp932')
    except: return s
def oxi_lines():
    raw=json.load(open(GLY,encoding='utf-8'))
    gl = raw if isinstance(raw,list) else (raw.get('glyphs') or [g for p in raw.get('pages',[]) for g in p.get('glyphs',[])])
    # glyph dump is flat in reading order; cluster by (page?,top). No page field -> use top only, in order.
    out=[]; cur=[]; lasttop=None
    for g in gl:
        c=dec(g['char']); top=g.get('top')
        if lasttop is None or abs(top-lasttop)<2:
            cur.append((c,g['x']))
        else:
            out.append(cur); cur=[(c,g['x'])]
        lasttop=top
    if cur: out.append(cur)
    return out

W=word_lines(); O=oxi_lines()
# char streams (skip spaces)
def stream(lines, isword):
    s=[]; meta=[]
    for li,ln in enumerate(lines):
        for j,item in enumerate(ln):
            c=item[0]
            if not c.strip(): continue
            s.append(c); meta.append((li,j))
    return s,meta
ws,wm=stream(W,True); os_,om=stream(O,False)
sm=difflib.SequenceMatcher(None,ws,os_,autojunk=False)
# For each Word line, collect matched (wj, oxi line, oxi idx)
wl_match={}   # word_line -> list of (wchar_x0,wchar_x1, oxi_line, oxi_x)
for blk in sm.get_matching_blocks():
    for k in range(blk.size):
        a=blk.a+k; b=blk.b+k
        wli,wj=wm[a]; oli,oj=om[b]
        wc,wx0,wx1=W[wli][wj]
        oc,ox=O[oli][oj]
        wl_match.setdefault(wli,[]).append((wx0,wx1,oli,ox,wc))
tot_drift=0; nlines=0; cls=Counter()
print(f"{'Wline':>5} {'n':>3} {'Wspan':>7} {'Ospan':>7} {'drift':>6}  text")
rows=[]
for wli in sorted(wl_match):
    m=wl_match[wli]
    if len(m)<6: continue
    # only chars on a SINGLE oxi line (the dominant one)
    olc=Counter(x[2] for x in m); dom_ol=olc.most_common(1)[0][0]
    mm=[x for x in m if x[2]==dom_ol]
    if len(mm)<6: continue
    wspan=mm[-1][1]-mm[0][0]
    ospan=mm[-1][3]-mm[0][3]   # oxi x of last - first (advance of last not included; approx)
    drift=wspan-ospan
    txt=''.join(x[4] for x in mm)
    rows.append((wli,len(mm),wspan,ospan,drift,txt))
    tot_drift+=drift; nlines+=1
rows.sort(key=lambda r:-r[4])
for wli,n,wspan,ospan,drift,txt in rows[:18]:
    print(f"{wli:>5} {n:>3} {wspan:>7.2f} {ospan:>7.2f} {drift:>6.2f}  {txt[:30]}")
print(f"\n  {nlines} lines, total drift (Word-Oxi span) = {tot_drift:.1f}pt, mean {tot_drift/max(1,nlines):.2f}pt/line")
print(f"  (drift>0 = Oxi narrower than Word for the same chars = Oxi can over-fit)")
