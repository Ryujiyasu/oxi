# -*- coding: utf-8 -*-
"""Compare Word vs Oxi line counts for 記載心得 #2 region to localize the
residual +1 (per-section column WIDTH causing extra wraps vs pure height)."""
import json,sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
W=json.load(open(r'pipeline_data/pagination_word/kyotei36spec.json',encoding='utf-8'))
# Word: paras 450..489, count lines via y-gap on same page (approx)
wp={p['i']:p for p in W['paragraphs']}
print('Word 記載心得#2 region (i, page, x, y):')
prev=None
for i in range(470,491):
    p=wp.get(i)
    if not p: continue
    g=''
    if prev and prev['page']==p['page']: g='gap=%.1f'%(p['y']-prev['y'])
    print('  i=%-3d p%d x=%5.1f y=%6.1f %s'%(i,p['page'],p['x'],p['y'],g))
    prev=p
# Oxi col width check: estimate from layout2 dump line x-extents on pages 4/5
d=json.load(open(r'C:/Users/ryuji/AppData/Local/Temp/kyotei_layout2.json',encoding='utf-8'))
for pgno in (3,4):
    p=d['pages'][pgno]
    txt=[e for e in p['elements'] if e['type']=='text' and e['text'].strip()]
    if not txt: 
        print('Oxi page %d empty'%(pgno+1)); continue
    xs=[e['x'] for e in txt]; rights=[e['x']+e['w'] for e in txt]
    # column detection: cluster x starts
    import statistics
    leftcol=[e for e in txt if e['x']<400]; rightcol=[e for e in txt if e['x']>=400]
    print('Oxi page %d: nText=%d  leftcol x_range=[%.0f,%.0f] rightcol x_range=[%.0f,%.0f]'%(
        pgno+1,len(txt),
        min((e['x'] for e in leftcol),default=0), max((e['x']+e['w'] for e in leftcol),default=0),
        min((e['x'] for e in rightcol),default=0), max((e['x']+e['w'] for e in rightcol),default=0)))
