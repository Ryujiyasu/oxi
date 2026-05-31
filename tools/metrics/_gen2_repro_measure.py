"""For each line-height repro: Word per-line Y gap (COM) vs Oxi (dump).
Single-line paras => gap = line_height(+after). Reveals Word's exact line
height and whether Oxi's 0.5pt quantization causes the per-para shortfall."""
import json, os, subprocess, tempfile, statistics
import win32com.client as win32

REPRO=r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\gen2_lineheight"
RENDER=r"C:\Users\ryuji\oxi-main\tools\oxi-gdi-renderer\target\release\oxi-gdi-renderer.exe"
VPOS=6;PAGE=3
NAMES=["cambria11_auto100_after0","cambria11_auto115_after0",
       "cambria11_auto115_after200","calibri11_auto115_after0"]

def word_gaps(docx):
    word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
    doc=word.Documents.Open(docx,ReadOnly=True);ys=[]
    try:
        for p in doc.Paragraphs:
            rng=p.Range;start=doc.Range(rng.Start,rng.Start)
            if start.Information(PAGE)!=1: continue
            if not p.Range.Text.strip(): continue
            ys.append(start.Information(VPOS))
    finally:
        doc.Close(False);word.Quit()
    return [round(ys[i+1]-ys[i],3) for i in range(len(ys)-1)]

def oxi_gaps(docx):
    with tempfile.TemporaryDirectory() as td:
        dump=os.path.join(td,"l.json")
        subprocess.run([RENDER,docx,os.path.join(td,"o"),"--dump-layout="+dump],
                       capture_output=True,timeout=60)
        d=json.load(open(dump,encoding="utf-8"))
    ys=[]
    for pg in d["pages"]:
        if pg["page"]!=1: continue
        # one y per para_idx (min y), body only
        seen={}
        for el in pg["elements"]:
            if el["type"]!="text" or not (el.get("text") or "").strip(): continue
            pi=el["para_idx"]
            if pi is None: continue
            seen[pi]=min(seen.get(pi,1e9),el["y"])
        ys=[seen[k] for k in sorted(seen)]
    return [round(ys[i+1]-ys[i],3) for i in range(len(ys)-1)]

print(f"{'variant':<30} {'word_gap':>10} {'oxi_gap':>10} {'diff/line':>10}")
print("-"*64)
for nm in NAMES:
    docx=os.path.join(REPRO,nm+".docx")
    wg=word_gaps(docx); og=oxi_gaps(docx)
    wm=statistics.median(wg) if wg else float('nan')
    om=statistics.median(og) if og else float('nan')
    print(f"{nm:<30} {wm:>10.3f} {om:>10.3f} {om-wm:>+10.3f}")
    print(f"{'  word gaps:':<30} {wg[:6]}")
    print(f"{'  oxi  gaps:':<30} {og[:6]}")
