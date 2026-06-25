#!/usr/bin/env python3
"""Derive Word's per-line cell 約物-compression (oikomi) from render-truth.

For the tokumei_08_01 charSpace=1453 family (de6e32/d4d126/6514/a1d6e4):
extract every Word-PDF text line, per-char advance, identify 約物 and their
compression (grid_pitch - advance), and classify the line as FULL (it wrapped
= packed to capacity) vs LAST (short). Report the compression distribution by
line-fullness + by 約物 type — the data to fit the compress-amount + the
oikomi/oidashi decision.

Usage: _cb_derive_family.py <word.pdf> [--fs 10.5] [--lines]
"""
import sys, statistics, collections

YAK_CLOSE = set('、。，．」』）〕】’”！？：；')   # closing/period/comma — Word compresses (right-aki)
YAK_OPEN  = set('「『（〔【‘“')                  # opening — left-aki only
YAK_MID   = set('・')
YAK = YAK_CLOSE | YAK_OPEN | YAK_MID

def is_full(c):
    o = ord(c)
    return o > 0x2000 and not (0xFF61 <= o <= 0xFF9F)

def cls(c):
    if c in YAK_CLOSE: return 'close'
    if c in YAK_OPEN:  return 'open'
    if c in YAK_MID:   return 'mid'
    return 'kanji'

def main():
    import fitz
    pdf = sys.argv[1]
    fs = float(sys.argv[sys.argv.index('--fs')+1]) if '--fs' in sys.argv else 10.5
    show = '--lines' in sys.argv
    doc = fitz.open(pdf)
    # gather lines: (page, ybase, xstart, xend, [(char, adv, size)])
    lines = []
    for pi, page in enumerate(doc):
        d = page.get_text('rawdict')
        rows = collections.defaultdict(list)
        for b in d['blocks']:
            for l in b.get('lines', []):
                for sp in l.get('spans', []):
                    sz = sp.get('size', 0)
                    for ch in sp.get('chars', []):
                        rows[(pi, round(ch['origin'][1], 0))].append((ch['origin'][0], ch['c'], sz))
        for key, chs in rows.items():
            chs.sort()
            if len(chs) < 2: continue
            advs = []
            for i in range(len(chs)):
                a = chs[i+1][0] - chs[i][0] if i+1 < len(chs) else None
                advs.append((chs[i][1], a, chs[i][2], chs[i][0]))
            lines.append((key[0], key[1], chs[0][0], chs[-1][0], advs))
    # grid pitch per fs = mode of kanji advances
    kadv = collections.defaultdict(list)
    for _,_,_,_,advs in lines:
        for c,a,sz,x in advs:
            if a and is_full(c) and cls(c)=='kanji' and 0<a<sz*1.3:
                kadv[round(sz,1)].append(a)
    pitch = {sz: statistics.median(v) for sz,v in kadv.items() if len(v)>=10}
    print("grid pitch by fs:", {k:round(v,3) for k,v in sorted(pitch.items())})
    p = pitch.get(round(fs,1)) or 10.79
    # detect FULL lines: a line is "full" if another line follows it in the same
    # x-column band (xstart within 6pt) on the same page at ybase+ (8..20pt).
    by_page = collections.defaultdict(list)
    for ln in lines: by_page[ln[0]].append(ln)
    full_flag = {}
    for pg, lns in by_page.items():
        lns2 = sorted(lns, key=lambda l:(l[1]))
        for i,(pgi,y,xs,xe,advs) in enumerate(lns2):
            wrapped = any(abs(xs2-xs)<8 and 8<(y2-y)<22 for (_,y2,xs2,_,_) in lns2 if y2>y)
            full_flag[(pgi,y,xs)] = wrapped
    # collect 約物 compression on FULL lines vs LAST lines, by class, for fs
    comp = {'full':collections.defaultdict(list), 'last':collections.defaultdict(list)}
    linestat = []
    for pgi,y,xs,xe,advs in lines:
        # only this fs
        fadv = [(c,a,sz,x) for c,a,sz,x in advs if a and abs(sz-fs)<0.6 and is_full(c)]
        if len(fadv) < 3: continue
        full = full_flag.get((pgi,y,xs), False)
        tag = 'full' if full else 'last'
        ncomp = 0; tot=0.0
        for c,a,sz,x in fadv:
            k = cls(c)
            if k != 'kanji':
                d = p - a       # compression (positive = compressed)
                comp[tag][k].append(d)
                if d > 0.5: ncomp += 1; tot += d
        linestat.append((pgi,y,len(fadv),full,ncomp,round(tot,1)))
    def rep(tag):
        print(f"\n=== {tag} lines, 約物 advance compression (pitch {p:.2f} - adv) ===")
        for k in ['close','open','mid']:
            v = comp[tag][k]
            if not v: continue
            v2=sorted(v)
            ncompressed=sum(1 for d in v if d>0.5)
            print(f"  {k:6s} n={len(v):4d} compressed(>0.5pt)={ncompressed:4d} ({100*ncompressed//len(v)}%) "
                  f"med={statistics.median(v):+.2f} p90={v2[int(len(v)*0.9)]:+.2f} max={v2[-1]:+.2f}")
    rep('full'); rep('last')
    nf=sum(1 for s in linestat if s[3]); nl=len(linestat)-nf
    print(f"\nfull lines={nf} last lines={nl}")
    if show:
        for pgi,y,n,full,nc,tot in sorted(linestat)[:60]:
            print(f"  p{pgi+1} y{y:5.0f} n{n:2d} {'FULL' if full else 'last'} ncomp={nc} totcomp={tot}")

if __name__ == '__main__':
    main()
