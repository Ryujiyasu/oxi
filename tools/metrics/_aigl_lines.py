# Per-line baseline Y on aiguideline PDF page 2 + page 3, compute consecutive
# gaps to read the ACTUAL rendered line height (Zen Old Mincho 12pt).
import fitz
doc=fitz.open(r"C:\tmp\aigl.pdf")
for pno in (1,2):
    pg=doc[pno]
    dd=pg.get_text("dict")
    lines=[]
    for b in dd["blocks"]:
        if b.get("type")!=0: continue
        for ln in b["lines"]:
            spans=ln["spans"]
            if not spans: continue
            txt="".join(s["text"] for s in spans).strip()
            if not txt: continue
            # baseline ~ origin y of first span
            org=spans[0]["origin"]
            y0=min(s["bbox"][1] for s in spans)
            sz=spans[0]["size"]
            lines.append((round(org[1],2), round(y0,2), round(sz,2), txt[:20]))
    lines.sort()
    print(f"=== page {pno+1} : {len(lines)} lines ===")
    prev=None
    for org_y,y0,sz,txt in lines:
        gap = (org_y-prev) if prev is not None else 0
        print(f"  baseY={org_y:7.2f} y0={y0:7.2f} sz={sz} gap={gap:6.2f}  {txt!r}")
        prev=org_y
