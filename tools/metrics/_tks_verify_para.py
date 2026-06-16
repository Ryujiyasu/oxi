# -*- coding: utf-8 -*-
import json,sys,os,tempfile
sys.stdout.reconfigure(encoding='utf-8',errors='replace')
import fitz
needle='法定労働時間を超えて労働させた場合には'
# Word PDF lines
doc=fitz.open(os.path.join(tempfile.gettempdir(),'tks_truth.pdf'))
print("=== WORD PDF lines for the para ===")
found=False; wc=0
for pi in range(len(doc)):
    lines=[]
    for blk in doc[pi].get_text('dict')['blocks']:
        if blk.get('type',0)!=0: continue
        for ln in blk.get('lines',[]):
            t=''.join(s['text'] for s in ln['spans'])
            lines.append((min(s['bbox'][1] for s in ln['spans']),t))
    for idx,(y,t) in enumerate(lines):
        if needle in t:
            found=True
            # print this line + following until next para marker (gap or 番号)
            for j in range(idx, min(idx+6,len(lines))):
                print(f"  p{pi+1} y{lines[j][0]:.0f}: {lines[j][1][:46]}")
            break
    if found: break
# Oxi dump lines
d=json.load(open('C:/tmp/tks_base.json',encoding='utf-8'))
print("=== OXI dump lines for the para (raw text elements) ===")
for pg in d['pages']:
    els=[e for e in pg.get('elements',[]) if e.get('type')=='text']
    # find the element with the needle start
    txt_by_y={}
    for e in els:
        txt_by_y.setdefault(round(e['y'],0),[]).append((e['x'],e.get('text','')))
    hit=None
    for y in sorted(txt_by_y):
        line=''.join(t for _,t in sorted(txt_by_y[y]))
        if '法定労働時間を超えて' in line:
            hit=y; break
    if hit is not None:
        for y in sorted(txt_by_y):
            if hit<=y<hit+120:
                line=''.join(t for _,t in sorted(txt_by_y[y]))
                print(f"  p{pg['page']} y{y:.0f}: {line[:46]}")
        break
