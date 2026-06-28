import json, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
def norm(s): return s.replace("　","").replace(" ","").strip()
W=json.load(open("C:/tmp/tks_word_glyphs.json",encoding="utf-8"))
# Build Word: per page, the SET of normalized line-prefixes (for content lookup)
wpage_lines={}   # page -> list of first-12-char prefixes of every line
for pi,pg in enumerate(W["pages"]):
    ys=defaultdict(list)
    for g in pg["glyphs"]:
        if g["y"]>=88 and g["y"]<785: ys[round(g["y"],0)].append(g)
    lines=[]
    for y in sorted(ys):
        gg=sorted(ys[y],key=lambda g:g["x"]); lines.append(norm("".join(g["char"] for g in gg))[:14])
    wpage_lines[pi+1]=lines
# index: prefix -> word page (first occurrence)
pref2wpage={}
for p in sorted(wpage_lines):
    for ln in wpage_lines[p]:
        if len(ln)>=8 and ln not in pref2wpage: pref2wpage[ln]=p
# Oxi page-tops
O=json.load(open("C:/tmp/tks_dump0.json",encoding="utf-8"))
def otop(pg):
    rows=defaultdict(lambda:['',1e9])
    for e in pg["elements"]:
        if e.get("type")=="text" and e.get("text","").strip() and round(e["y"],1)>=88:
            y=round(e["y"],1); rows[y][0]+=e["text"]
    if not rows: return ""
    y0=min(rows); return norm(rows[y0][0])[:14]
ahead=behind=match=0; rows=[]
for pg in O["pages"]:
    p=pg["page"]; ot=otop(pg)
    wp=pref2wpage.get(ot)
    if wp is None:
        # try partial
        for k,v in pref2wpage.items():
            if ot and len(ot)>=10 and ot[:10]==k[:10]: wp=v; break
    if wp is None: rows.append((p,'?',ot)); continue
    delta=p-wp
    if delta==0: match+=1
    elif delta<0: ahead+=1; rows.append((p,wp,ot))   # Oxi ahead (−1)
    else: behind+=1; rows.append((p,wp,ot))          # Oxi behind (+1)
print(f"match(delta0)={match} ahead(−1)={ahead} behind(+1)={behind} unknown={sum(1 for r in rows if r[1]=='?')}")
print("\n-- Oxi BEHIND (+1: Oxi page p shows content from Word page wp<p) --")
for p,wp,ot in rows:
    if wp!='?' and isinstance(wp,int) and p>wp: print(f"  Oxi p{p:3} = Word p{wp:3} (Δ+{p-wp}) | {ot}")
print("\n-- Oxi AHEAD (−1) --")
for p,wp,ot in rows:
    if wp!='?' and isinstance(wp,int) and p<wp: print(f"  Oxi p{p:3} = Word p{wp:3} (Δ{p-wp}) | {ot}")
