"""Reconstruct per-paragraph (y, fulltext) for Word and Oxi, align by text,
print Y and consecutive gap side-by-side to localize the exact drift source."""
import json
import win32com.client as win32
DOCX=r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\gen2_060_Employee_Agreement.docx"
DUMP=r"C:/tmp/gen2dump/layout.json"
VPOS=6;PAGE=3

# Oxi: group elements by (para_idx, row, col); paragraph y=min, text=concat in x order
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
    els.sort(key=lambda e:e["x"])
    txt="".join(e.get("text") or "" for e in els).strip()
    if not txt: continue
    y=min(e["y"] for e in els)
    oxi.append((y,txt))
oxi.sort()

word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
doc=word.Documents.Open(DOCX,ReadOnly=True)
wseq=[]
try:
    for p in doc.Paragraphs:
        rng=p.Range;start=doc.Range(rng.Start,rng.Start)
        if start.Information(PAGE)!=1: continue
        t=p.Range.Text.strip()
        if not t: continue
        wseq.append((round(start.Information(VPOS),2),t))
finally:
    doc.Close(False);word.Quit()

# Align by exact text match, greedy from start with skip tolerance
def norm(t): return t[:18]
oi=0
print(f"{'word_y':>7} {'oxi_y':>7} {'diff':>6} | {'wgap':>5} {'ogap':>5} {'dg':>5}  text")
print("-"*72)
pw=po=None
for wy,wt in wseq:
    # find next oxi with matching text at/after oi
    j=oi
    while j<len(oxi) and norm(oxi[j][1])!=norm(wt): j+=1
    if j<len(oxi):
        oy,ot=oxi[j]; oi=j+1
        wgap=(wy-pw) if pw is not None else 0
        ogap=(oy-po) if po is not None else 0
        dg=ogap-wgap
        m=" <==" if abs(dg)>0.6 and pw is not None else ""
        print(f"{wy:>7.1f} {oy:>7.1f} {oy-wy:>+6.1f} | {wgap:>5.1f} {ogap:>5.1f} {dg:>+5.1f}  {wt[:26]}{m}")
        pw,po=wy,oy
    else:
        print(f"{wy:>7.1f} {'--':>7}            (no oxi) {wt[:26]}")
