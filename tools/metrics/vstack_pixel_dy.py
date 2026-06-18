# -*- coding: utf-8 -*-
"""Per-line pixel dy: Word PDF rendered baseline vs Oxi --dump-glyphs baseline.
The PIXEL truth for the vertical-stack re-derivation (NOT COM box tops).
Aligns lines per page by y-order, reports dy(line)=Oxi-Word and the per-page
first-line offset + accumulation slope.
Usage: python vstack_dy.py <docx_basename_glob>   (e.g. gen2_005)
"""
import os, sys, json, subprocess, glob
import fitz
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
ROOT=r'C:/Users/ryuji/oxi-main'
DW=os.path.join(ROOT,'tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
DOCDIR=os.path.join(ROOT,'tools/golden-test/documents/docx')
base=sys.argv[1]
cands=[p for p in glob.glob(os.path.join(DOCDIR,base+'*.docx')) if not os.path.basename(p).startswith('~$')]
docx=cands[0]; stem=base
pdf=f'C:/tmp/ta/{stem}_vw.pdf'; gj=f'C:/tmp/ta/{stem}_vg.json'

# Word PDF render-truth
if not os.path.exists(pdf) or '--regen' in sys.argv:
    import win32com.client, pythoncom
    pythoncom.CoInitialize(); w=win32com.client.DispatchEx('Word.Application'); w.Visible=False
    try:
        d=w.Documents.Open(os.path.abspath(docx),ReadOnly=True); d.ExportAsFixedFormat(pdf,17); d.Close(False)
    finally: w.Quit()

def word_lines(pdf):
    doc=fitz.open(pdf); pages=[]
    for pg in doc:
        chs=[]
        for blk in pg.get_text('rawdict').get('blocks',[]):
            for ln in blk.get('lines',[]):
                for sp in ln.get('spans',[]):
                    for c in sp.get('chars',[]):
                        if c['c'].strip(): chs.append((c['origin'][1],c['origin'][0],c['c'],sp.get('size')))
        chs.sort(key=lambda c:(round(c[0]),c[1])); lines=[]
        for c in chs:
            if lines and abs(c[0]-lines[-1][-1][0])<2.5: lines[-1].append(c)
            else: lines.append([c])
        pages.append([(min(x[0] for x in ln), min(x[1] for x in ln), ln[0][3], "".join(x[2] for x in ln)) for ln in lines])
    return pages

subprocess.run([DW,docx,f'C:/tmp/ta/{stem}_v',f'--dump-glyphs={gj}','150'],capture_output=True)
g=json.load(open(gj,encoding='utf-8'))
def oxi_lines(g):
    pages=[]
    for pg in g['pages']:
        gl=[x for x in pg['glyphs'] if x['char'].strip()]
        gl.sort(key=lambda c:(round(c['baseline']),c['x'])); lines=[]
        for c in gl:
            if lines and abs(c['baseline']-lines[-1][-1]['baseline'])<2.5: lines[-1].append(c)
            else: lines.append([c])
        pages.append([(min(x['baseline'] for x in ln), min(x['x'] for x in ln), ln[0]['font_size'], "".join(x['char'] for x in ln)) for ln in lines])
    return pages

W=word_lines(pdf); O=oxi_lines(g)
print(f"{stem}: Word pages={len(W)} Oxi pages={len(O)}")
for pi in range(min(len(W),len(O))):
    wl=W[pi]; ol=O[pi]; n=min(len(wl),len(ol))
    print(f"--- page {pi+1}: Word {len(wl)} lines, Oxi {len(ol)} lines ---")
    dys=[]
    for i in range(n):
        wb,wx,wfs,wt=wl[i]; ob,ox,ofs,ot=ol[i]
        dy=ob-wb; dys.append(dy)
        if i<4 or i>=n-3 or abs(dy)>1.5 or ofs>=18:
            print(f"  L{i:2d} fs={ofs:4.1f} Wbase={wb:7.2f} Obase={ob:7.2f} dy={dy:+5.2f}  W={wt[:8]!r}")
    if dys:
        print(f"  page dy: first={dys[0]:+.2f} last={dys[-1]:+.2f} mean={sum(dys)/len(dys):+.2f} span={max(dys)-min(dys):.2f}")
