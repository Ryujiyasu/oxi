# -*- coding: utf-8 -*-
"""Compare Word vs Oxi per-paragraph TOP-Y for harassmanual paras 10-40 to
localize the +1x3 height over-count (post-S567). Word via COM (R30 collapsed
start), Oxi via GDI layout dump (para_idx -> min y on its page)."""
import os, io, json, collections
import win32com.client as win32
DOCX = os.path.abspath("tools/golden-test/documents/docx/harassmanual_001466344.docx")
DUMP = r"C:\tmp\hm_fix.json"
# --- Oxi: para_idx -> (page, min_y, first_text) for body text ---
d = json.load(io.open(DUMP, encoding="utf-8"))
oxi = {}  # para_idx -> [page, miny, text@miny]
for pi, pg in enumerate(d["pages"], 1):
    for e in pg["elements"]:
        if e.get("type") == "text" and e.get("text"):
            k = e.get("para_idx")
            if k is None: continue
            y = e["y"]
            if k not in oxi or y < oxi[k][1]:
                oxi[k] = [pi, y, e["text"]]
# collect first-line text per para (concatenate elems at miny)
oxi_txt = collections.defaultdict(list)
for pi, pg in enumerate(d["pages"], 1):
    for e in pg["elements"]:
        if e.get("type")=="text" and e.get("text"):
            k=e.get("para_idx")
            if k in oxi and abs(e["y"]-oxi[k][1])<0.3 and oxi[k][0]==pi:
                oxi_txt[k].append((e["x"], e["text"]))
for k in oxi_txt: oxi[k][2] = "".join(t for _,t in sorted(oxi_txt[k]))
# --- Word: per-para page + top Y ---
word = win32.gencache.EnsureDispatch("Word.Application"); word.Visible=False
doc = word.Documents.Open(DOCX, ReadOnly=True)
wrows=[]
try:
    for i in range(1, min(doc.Paragraphs.Count,42)+1):
        rng=doc.Paragraphs(i).Range; col=doc.Range(rng.Start,rng.Start)
        wrows.append((i, col.Information(3), col.Information(6), rng.Text.strip()[:16]))
finally:
    doc.Close(False); word.Quit()
# --- align by text prefix ---
print("Word: i page  topY   dY   | Oxi page topY   | text")
oxi_by_text={ v[2][:10]: (k,v) for k,v in oxi.items() if v[2] }
prev_wy=None; prev_wp=None
for (i,wp,wy,wt) in wrows:
    dy = (wy-prev_wy) if (prev_wy is not None and wp==prev_wp) else None
    op=oy="-"
    key=wt[:10]
    if key and key in oxi_by_text:
        _,v=oxi_by_text[key]; op=v[0]; oy="%.1f"%v[1]
    dys = "%5.1f"%dy if dy is not None else "  -- "
    div = "  <<<" if (op!="-" and op!=wp) else ""
    print(f"{i:3d} p{wp} {wy:6.1f} {dys} | p{op} {oy:>6} | {wt!r}{div}")
    prev_wy=wy; prev_wp=wp
