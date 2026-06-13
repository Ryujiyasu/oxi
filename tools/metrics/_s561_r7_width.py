# -*- coding: utf-8 -*-
"""S561 DEFINITIVE: roudoujoken r7 content cell (gridSpan=2) wrap-budget is ~14pt
too wide. Content-cell box border L=123.7 R=555.6 (width 431.9pt = tcW 8640tw),
inner SHOULD be 431.9-2.6-2.6=426.8 (L/R cellMar 52tw). But the (5)裁量 line's
text paints to x=567.2 -- 11.6pt PAST the right border -- so Oxi wraps at ~441pt,
fitting (5)裁量 (~420pt) on 1 line where Word's 426.8pt budget wraps 労働者… to
line 2. That 1 line is roudoujoken's -1. Re-run after rendering the dump."""
import json,sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
d=json.load(open(r'C:/Users/ryuji/AppData/Local/Temp/roudou_layout.json',encoding='utf-8'))
p1=d['pages'][0]
vb=sorted(set(round(e['x'],1) for e in p1['elements'] if e['type']=='border' and e['h']>0 and 350<e['y']<620))
print('r7-band vertical borders x:', vb, '(content cell L=col0/col1 bound, R=table right)')
lines=defaultdict(list)
for e in p1['elements']:
    if e['type']=='text' and e['text'].strip() and e.get('cell_row_idx')==7:
        lines[round(e['y'])].append(e)
for y,es in sorted(lines.items()):
    es=sorted(es,key=lambda e:e['x']); txt=''.join(e['text'] for e in es)
    if '裁量' in txt:
        L=es[0]['x']; R=max(e['x']+e['w'] for e in es)
        print('(5)裁量 line: x_left=%.1f x_right=%.1f width=%.1f  (cell right border=555.6 -> overshoot=%.1f)'%(L,R,R-L,R-555.6))
