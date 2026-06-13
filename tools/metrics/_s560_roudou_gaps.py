# -*- coding: utf-8 -*-
import json,sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
d=json.load(open(r'C:/Users/ryuji/AppData/Local/Temp/roudou_layout.json',encoding='utf-8'))
# para_idx -> (page, min_y, max_y, text)
rec=defaultdict(lambda:[None,1e9,-1e9,None])
for pgno,p in enumerate(d['pages']):
    for e in p['elements']:
        if e['type']!='text' or not e['text'].strip(): continue
        pi=e.get('para_idx')
        if pi is None: continue
        r=rec[pi]
        if e['y']<r[1]: r[1]=e['y']; 
        r[2]=max(r[2],e['y']); 
        if r[0] is None: r[0]=pgno+1
        if r[3] is None: r[3]=e['text'][:18]
print('Oxi roudoujoken: para_idx -> page, top_y, text  (around the 8.「休暇」 region)')
prev=None
for pi in sorted(rec):
    r=rec[pi]
    if r[3] and ('休暇' in r[3] or '年次' in r[3] or pi>=0):
        pass
    # print all in a window; find the 8. para
for pi in sorted(rec):
    r=rec[pi]
    g=''
    if prev and prev[0]==r[0]: g='gap=%.1f'%(r[1]-prev[1])
    flag=' <<<' if (r[3] and ('８' in r[3] or '休暇' in r[3])) else ''
    # only print pages 3-4 region
    if r[0] in (3,4):
        print('  pi=%-3d p%d y=%6.1f  %s%s%s'%(pi,r[0],r[1],r[3],g,flag))
    prev=r
