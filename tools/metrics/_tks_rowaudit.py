# -*- coding: utf-8 -*-
import json, fitz, sys
io=sys.stdout; io.reconfigure(encoding='utf-8',errors='replace')
PDF=r'C:\Users\ryuji\AppData\Local\Temp\tks_truth.pdf'
DUMP=r'C:/tmp/tks_dump.json'  # Oxi S586 dump (89 pages)

wd=fitz.open(PDF)
ox=json.load(open(DUMP,encoding='utf-8'))

def word_hborders(pidx):
    """horizontal line y-positions on Word PDF page pidx (0-based)."""
    pg=wd[pidx]; ys=[]
    for dr in pg.get_drawings():
        for it in dr['items']:
            if it[0]=='l':
                p1,p2=it[1],it[2]
                if abs(p1.y-p2.y)<0.6 and abs(p1.x-p2.x)>5:
                    ys.append(round(p1.y,1))
            elif it[0]=='re':
                r=it[1]; ys.append(round(r.y0,1)); ys.append(round(r.y1,1))
    return sorted(set(ys))

def oxi_hborders(p):
    ys=[e['y'] for e in p['elements'] if e.get('type')=='border' and abs(e.get('h',0))<0.6 and e.get('w',0)>5]
    return sorted(set(round(y,1) for y in ys))

# table extent per page = sum of gaps between consecutive h-borders that are <some max (a row)
def extent(ys):
    if len(ys)<2: return 0.0, 0
    span=ys[-1]-ys[0]
    return span, len(ys)

print('PAGE | Word: nHB span | Oxi: nHB span   (Word pidx vs Oxi pidx)')
# 賃金 chapter: Word p46-64 (idx45-63), Oxi p46-63 (idx45-62) under S586 (-1)
wtot=0; otot=0
for wp in range(45,64):
    ws,wn=extent(word_hborders(wp))
    op=wp-1  # Oxi 1 page ahead in chapter
    if 0<=op<len(ox['pages']):
        os_,on=extent(oxi_hborders(ox['pages'][op]))
    else: os_,on=0,0
    wtot+=ws; otot+=os_
    print(f'Wp{wp+1:>2}/Op{op+1:>2} | {wn:>3} {ws:7.1f} | {on:>3} {os_:7.1f}')
print(f'TOTAL table h-border span  Word {wtot:.0f}  Oxi {otot:.0f}  diff {wtot-otot:.0f}')
