import json,sys
import win32com.client as win32
NAME=sys.argv[1]; DUMP=sys.argv[2]
DOCX=fr"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\{NAME}.docx"
VPOS=6;PAGE=3
d=json.load(open(DUMP,encoding="utf-8"))
groups={}
for pg in d["pages"]:
    if pg["page"]!=1: continue
    for el in pg["elements"]:
        if el["type"]!="text": continue
        k=(el["para_idx"],el.get("cell_row_idx"),el.get("cell_col_idx"))
        groups.setdefault(k,[]).append(el)
oxi=[]
for k,els in groups.items():
    els.sort(key=lambda e:e["x"]); txt="".join(e.get("text") or "" for e in els).strip()
    if txt: oxi.append((min(e["y"] for e in els),txt))
oxi.sort()
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
doc=word.Documents.Open(DOCX,ReadOnly=True);wseq=[]
try:
    for p in doc.Paragraphs:
        rng=p.Range;start=doc.Range(rng.Start,rng.Start)
        if start.Information(PAGE)!=1: continue
        t=p.Range.Text.strip()
        if t: wseq.append((round(start.Information(VPOS),2),t))
finally:
    doc.Close(False);word.Quit()
norm=lambda t:t[:18]; oi=0
print(f"{'word_y':>7} {'oxi_y':>7} {'diff':>6}  text")
for wy,wt in wseq:
    if wt.startswith("Cell "): continue
    j=oi
    while j<len(oxi) and norm(oxi[j][1])!=norm(wt): j+=1
    if j<len(oxi):
        oy=oxi[j][0]; oi=j+1
        print(f"{wy:>7.1f} {oy:>7.1f} {oy-wy:>+6.1f}  {wt[:30]}")
