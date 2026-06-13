# -*- coding: utf-8 -*-
"""Word vs Oxi line counts for ALL roudoujoken page-3 paras (find the ~14pt
under-count that lets Oxi fit ８「休暇」(167) on p3 where Word pushes to p4)."""
import sys,os,json,subprocess,tempfile
import win32com.client as win32
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
R=os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
docx=os.path.abspath('tools/golden-test/documents/docx/roudoujoken_001161383.docx')
# Oxi dump
with tempfile.TemporaryDirectory() as td:
    dj=os.path.join(td,'l.json')
    subprocess.run([R,docx,os.path.join(td,'p'),'150','--dump-layout='+dj],capture_output=True)
    d=json.load(open(dj,encoding='utf-8'))
# Oxi page 3 (index 2): para_idx -> n lines + text
oxi=defaultdict(set); otxt={}
for e in d['pages'][2]['elements']:
    if e['type']=='text' and e['text'].strip():
        pi=e.get('para_idx'); oxi[pi].add(round(e['y'])); otxt.setdefault(pi,'')
        if len(otxt[pi])<14: otxt[pi]+=e['text']
# Word: page-3 paras line counts
wd=win32.gencache.EnsureDispatch('Word.Application'); wd.Visible=False
doc=wd.Documents.Open(docx,ReadOnly=True)
def dec(s):
    try: return s.encode('latin1').decode('utf-8')
    except: return s
print('Word page-3 paras (i, Wlines, text) + Oxi match:')
def norm(s): return ''.join((s or '').split())[:10]
oxi_by_t={}
for pi,t in otxt.items(): oxi_by_t.setdefault(norm(t),(len(oxi[pi]),pi))
for i in range(1,doc.Paragraphs.Count+1):
    rng=doc.Paragraphs(i).Range
    if rng.Information(3)!=3: continue
    txt=rng.Text.strip()
    if not txt: continue
    ys=set()
    for k in range(0,len(txt)):
        ch=doc.Range(rng.Start+k,rng.Start+k+1)
        try: ys.add(round(ch.Information(6)*2)/2)
        except: pass
    sy=sorted(ys); wl=0; prev=-99
    for y in sy:
        if y-prev>5: wl+=1; prev=y
    om=oxi_by_t.get(norm(txt))
    flag=''
    if om and om[0]!=wl: flag=' <<< MISMATCH'
    print('  i=%-3d W=%d O=%s %r%s'%(i, wl, om[0] if om else '?', txt[:22], flag))
doc.Close(False); wd.Quit()
