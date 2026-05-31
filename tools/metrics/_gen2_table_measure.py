"""Table block height: Word (BEFORE_MARK -> AFTER_MARK Y gap) vs Oxi dump.
Isolates whether Oxi's 5-row table total height is short vs Word."""
import json, os, subprocess, tempfile
import win32com.client as win32
DOCX=r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\gen2_lineheight\table_5x3_cambria11.docx"
RENDER=r"C:\Users\ryuji\oxi-main\tools\oxi-gdi-renderer\target\release\oxi-gdi-renderer.exe"
VPOS=6;PAGE=3
# Word
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
doc=word.Documents.Open(DOCX,ReadOnly=True)
wb=wa=None; rowYs=[]
try:
    for p in doc.Paragraphs:
        rng=p.Range;start=doc.Range(rng.Start,rng.Start)
        y=start.Information(VPOS);t=p.Range.Text.strip()
        if t=="BEFORE_MARK": wb=y
        elif t=="AFTER_MARK": wa=y
        elif t.startswith("Cell ") and t.endswith("-1"):
            rowYs.append(round(y,2))
finally:
    doc.Close(False);word.Quit()
print(f"WORD: before={wb:.2f} after={wa:.2f}  table_block={wa-wb:.2f}pt")
print(f"WORD row-1 cell Ys: {rowYs}  row pitches: {[round(rowYs[i+1]-rowYs[i],2) for i in range(len(rowYs)-1)]}")
# Oxi
with tempfile.TemporaryDirectory() as td:
    dump=os.path.join(td,"l.json")
    subprocess.run([RENDER,DOCX,os.path.join(td,"o"),"--dump-layout="+dump],capture_output=True,timeout=60)
    d=json.load(open(dump,encoding="utf-8"))
ob=oa=None; orow={}
for pg in d["pages"]:
    if pg["page"]!=1: continue
    for el in pg["elements"]:
        if el["type"]!="text": continue
        t=(el.get("text") or "").strip()
        if t=="BEFORE_MARK": ob=el["y"]
        elif t=="AFTER_MARK": oa=el["y"]
        elif t.startswith("Cell ") and el.get("cell_col_idx")==0:
            r=el.get("cell_row_idx"); orow[r]=min(orow.get(r,1e9),el["y"])
orows=[round(orow[k],2) for k in sorted(orow)]
print(f"OXI : before={ob:.2f} after={oa:.2f}  table_block={oa-ob:.2f}pt")
print(f"OXI row-1 cell Ys: {orows}  row pitches: {[round(orows[i+1]-orows[i],2) for i in range(len(orows)-1)]}")
print(f"=> table block diff (oxi-word): {(oa-ob)-(wa-wb):+.2f}pt over 5 rows")
