"""Drift curve via UNIQUE anchor paragraphs (Word COM vs Oxi dump), gen2_060 p1."""
import json
import win32com.client as win32
DOCX=r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\gen2_060_Employee_Agreement.docx"
DUMP=r"C:/tmp/gen2dump/layout.json"
VPOS=6;PAGE=3
ANCHORS=["Employee Agreement","Overview","The following report","Purpose",
"Based on our assessment","This policy is effective","Please do not hesitate",
"Background","This document has been","Details"]

# Oxi: first text element whose text startswith anchor
d=json.load(open(DUMP,encoding="utf-8"))
els=[el for pg in d["pages"] if pg["page"]==1 for el in pg["elements"]
     if el["type"]=="text" and (el.get("text") or "").strip()]
def oxi_y(a):
    cand=[el["y"] for el in els if el["text"].strip().startswith(a[:14])]
    return min(cand) if cand else None

word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
doc=word.Documents.Open(DOCX,ReadOnly=True)
wy={}
try:
    for p in doc.Paragraphs:
        rng=p.Range;start=doc.Range(rng.Start,rng.Start)
        if start.Information(PAGE)!=1: continue
        t=p.Range.Text.strip()
        for a in ANCHORS:
            if t.startswith(a[:14]) and a not in wy:
                wy[a]=round(start.Information(VPOS),2)
finally:
    doc.Close(False);word.Quit()

print(f"{'word_y':>8} {'oxi_y':>8} {'diff':>6} {'step':>6}  anchor")
print("-"*58)
prev=None
for a in ANCHORS:
    w=wy.get(a);o=oxi_y(a)
    if w is None or o is None:
        print(f"{str(w):>8} {str(o):>8}   --   --   {a}  (missing)");continue
    diff=o-w; step=(diff-prev) if prev is not None else 0.0
    print(f"{w:>8.1f} {o:>8.1f} {diff:>+6.1f} {step:>+6.1f}  {a}")
    prev=diff
