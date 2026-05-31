"""Per-paragraph first-line Y matched BY TEXT (para_idx differs: Word counts
table cells as paragraphs). Localizes Oxi vertical drift vs Word on gen2_060 p1."""
import json
import win32com.client as win32

DOCX = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\gen2_060_Employee_Agreement.docx"
DUMP = r"C:/tmp/gen2dump/layout.json"
VPOS = 6; PAGE = 3

# Oxi: ordered (y, text) first element per (para_idx, cell) on page 1
d = json.load(open(DUMP, encoding="utf-8"))
seen = {}
for pg in d["pages"]:
    if pg["page"] != 1: continue
    for el in pg["elements"]:
        if el["type"] != "text" or not (el.get("text") or "").strip(): continue
        key = (el["para_idx"], el.get("cell_row_idx"), el.get("cell_col_idx"))
        y = el["y"]
        if key not in seen or y < seen[key][0]:
            seen[key] = (y, el["text"].strip())
oxi_seq = sorted(seen.values())  # by y

# Word: ordered paragraphs page 1 with text + Y
word = win32.gencache.EnsureDispatch("Word.Application"); word.Visible=False
doc = word.Documents.Open(DOCX, ReadOnly=True)
word_seq=[]
try:
    for p in doc.Paragraphs:
        rng=p.Range; start=doc.Range(rng.Start,rng.Start)
        if start.Information(PAGE)!=1: continue
        t=p.Range.Text.strip()
        if not t: continue
        word_seq.append((round(start.Information(VPOS),2), t[:24]))
finally:
    doc.Close(False); word.Quit()

# match distinctive (non "Cell") paragraphs by text, in order
def distinctive(t): return not t.startswith("Cell ")
W=[(y,t) for y,t in word_seq if distinctive(t)]
O=[(y,t[:24]) for y,t in oxi_seq if distinctive(t)]
print("Word distinctive paras:",len(W)," Oxi:",len(O))
print(f"{'word_y':>8} {'oxi_y':>8} {'diff':>6} {'step':>6}  text")
print("-"*64)
prev=0.0
n=min(len(W),len(O))
for i in range(n):
    wy,wt=W[i]; oy,ot=O[i]
    diff=oy-wy; step=diff-prev
    flag=" <==" if abs(step)>0.8 else ""
    tag="" if wt[:12]==ot[:12] else f" [W:{wt[:14]}|O:{ot[:14]}]"
    print(f"{wy:>8.1f} {oy:>8.1f} {diff:>+6.1f} {step:>+6.1f}  {wt[:20]}{tag}{flag}")
    prev=diff
