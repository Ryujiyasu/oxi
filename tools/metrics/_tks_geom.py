# -*- coding: utf-8 -*-
import os, sys, json, tempfile, fitz, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
TMP=tempfile.gettempdir()
PDF=os.path.join(TMP,'tks_truth.pdf')
DUMP=os.environ.get('TKS_DUMP', os.path.join(TMP,'tks_s586.json'))
isfoot=lambda t: bool(re.fullmatch(r'\s*\d{1,3}\s*', t))
def word_pg(pno):
    pg=fitz.open(PDF)[pno-1]; rd=pg.get_text('rawdict'); chars=[]
    for blk in rd['blocks']:
        if blk.get('type',0)!=0: continue
        for ln in blk.get('lines',[]):
            for sp in ln['spans']:
                for ch in sp['chars']:
                    b=ch['bbox']; y0=b[1]
                    if y0<60 or y0>800: continue
                    chars.append((ch['c'],(b[1]+b[3])/2))
    rows={}
    for c,yc in chars: rows.setdefault(round(yc/3.0),[]).append((c,yc))
    ys=[]
    for k in sorted(rows):
        txt=''.join(c for c,_ in rows[k])
        if isfoot(txt): continue
        ys.append(round(sum(y for _,y in rows[k])/len(rows[k]),1))
    return ys
def oxi_pg(d,pno):
    for pg in d['pages']:
        if pg['page']!=pno: continue
        rows={}
        for el in pg.get('elements',[]):
            if el.get('type')!='text' or not el.get('text'): continue
            rows.setdefault(round(el['y'],0),[]).append(el)
        ys=[]
        for k in sorted(rows):
            txt=''.join(e.get('text','') for e in rows[k])
            if isfoot(txt): continue
            ys.append(k)
        return ys
    return []
d=json.load(open(DUMP,encoding='utf-8'))
print(f"{'Wp':>4} {'Wn':>3} {'Wy0':>6} {'Wy1':>6} | {'Op':>3} {'On':>3} {'Oy0':>6} {'Oy1':>6}")
for wp in range(46,60):
    wy=word_pg(wp)
    op=wp  # under S586 chapter is aligned Wp46-58 so Oxi page == Word page here
    oy=oxi_pg(d,op)
    if not wy or not oy: continue
    print(f"{wp:>4} {len(wy):>3} {wy[0]:>6} {wy[-1]:>6} | {op:>3} {len(oy):>3} {oy[0]:>6} {oy[-1]:>6}")
