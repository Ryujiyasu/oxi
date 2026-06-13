# -*- coding: utf-8 -*-
"""S562: per-line over-pack of roudoujoken page-3 記載要領 paras (Word vs Oxi).
Word: count lines per para from per-char Information(6) y-grouping (COM).
Oxi: from the dump, group para's text elements by y. Compare line counts."""
import sys,os,json
import win32com.client as win32
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
# Oxi line counts per body para_idx (page 3 = index 2), for the 記載要領 paras
d=json.load(open(r'C:/Users/ryuji/AppData/Local/Temp/rd_dbg2.json',encoding='utf-8'))
oxi=defaultdict(set)
oxitext={}
for pgno in (2,3):
    for e in d['pages'][pgno]['elements']:
        if e['type']=='text' and e['text'].strip():
            pi=e.get('para_idx')
            if pi is None: continue
            oxi[pi].add(round(e['y']))
            oxitext.setdefault(pi,'')
            if len(oxitext[pi])<14: oxitext[pi]+=e['text']
# Word: open doc, for body paras on page 3 (記載要領), count lines via Information(6)
wd=win32.gencache.EnsureDispatch('Word.Application'); wd.Visible=False
doc=wd.Documents.Open(os.path.abspath('tools/golden-test/documents/docx/roudoujoken_001161383.docx'),ReadOnly=True)
print('Word vs Oxi line counts for page-3 記載要領 paras:')
targets=['裁量労働制：','６．「始業','７．「休日','事業場外みなし','交替制：']
for i in range(1,doc.Paragraphs.Count+1):
    rng=doc.Paragraphs(i).Range; txt=rng.Text.strip()
    if not any(t in txt for t in targets): continue
    if rng.Information(3)!=3: continue  # page 3 only
    # count lines via per-char y
    ys=set()
    for k in range(0,len(txt),1):
        ch=doc.Range(rng.Start+k,rng.Start+k+1)
        try: ys.add(round(ch.Information(6)*2)/2)
        except: pass
    # cluster ys into lines (within 3pt)
    sy=sorted(ys); lines=0; prev=-99
    for y in sy:
        if y-prev>5: lines+=1; prev=y
    print('  Word %d lines  %r'%(lines, txt[:26]))
doc.Close(False); wd.Quit()
print('\nOxi page-3 paras (para_idx: n_lines  text):')
for pi in sorted(oxi):
    if any(t.replace('：','') in oxitext[pi] or '裁量' in oxitext[pi] or '始業' in oxitext[pi] for t in ['裁量','始業']):
        print('  Oxi pi=%-3d %d lines  %r'%(pi,len(oxi[pi]),oxitext[pi]))
