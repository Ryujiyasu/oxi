"""python tools/metrics/_pb_c14wpj_gen.py   (builds the docx, measures it in a fresh Word, prints the flips)

New-Word wpJustification allowance: does it scale with the gap count (G) or with
the line width?  Host settings KEPT (the real doc's compat block), jc=both, one
'mmmm'*G + 'final.' line per page; sweep R (room left for the candidate) in 0.3pt
steps.  Also a width axis: G=9 with a LEFT indent L (narrower column, same gaps)."""
import json, os, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")
import _pb_c14floor_gen as P
from lxml import etree

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "pipeline_data", "_pb_c14wpj_newpc")
SPECS = []
for G in (4, 6, 9, 12):
    for k in range(0, 41):
        R = 30.0 + 0.6 * k - (G - 9) * 1.5          # centred on a ~1.5/gap guess
        ri = P.ri_for_R(G, R)
        if ri < 0:
            continue
        SPECS.append((f"G{G}_R{int(round(R * 100)):05d}", P.line_text(G, "final."),
                      {"G": G, "R": round(R, 2), "L": 0, "ri": ri, "cand": "final."}))
for L in (36, 72):                                  # narrower column, G=9
    for k in range(0, 41):
        R = 30.0 + 0.6 * k
        # keep the room R for the candidate: right indent shrinks by L
        ri = P.ri_for_R(9, R) - L * 20
        if ri < 0:
            continue
        SPECS.append((f"L{L}_R{int(round(R * 100)):05d}", P.line_text(9, "final."),
                      {"G": 9, "R": round(R, 2), "L": L, "ri": ri, "cand": "final."}))

_add = P.add_para
def add_para(body, name, text, ri_tw, bid):
    _add(body, name, text, ri_tw, bid)
    L = META[name]["L"]
    if L:
        ind = body[-1].find("w:pPr/w:ind", P.NS)
        ind.set(P.qn("left"), str(L * 20))
META = {n: m for n, _, m in SPECS}
P.add_para = add_para
P.specimens = lambda: SPECS

os.makedirs(OUT, exist_ok=True)
with zipfile.ZipFile(P.HOST) as z:
    parts = {n: z.read(n) for n in z.namelist()}
parts["word/document.xml"] = P.rewrite_document(parts["word/document.xml"], SPECS)
with zipfile.ZipFile(os.path.join(OUT, "floor.docx"), "w", zipfile.ZIP_DEFLATED) as z:
    for n, b in parts.items():
        z.writestr(n, b)
P.OUTDIR = OUT
try:
    P.measure()
except Exception as e:
    print("summary:", e)
res = json.load(open(os.path.join(OUT, "_result.json")))
for key in [("G", g) for g in (4, 6, 9, 12)] + [("L", l) for l in (36, 72)]:
    rows = sorted((r for n, r in res.items() if (n.startswith(f"G{key[1]}_") if key[0] == "G" else n.startswith(f"L{key[1]}_"))),
                  key=lambda r: r["R"])
    keep = [r["R"] for r in rows if r["com_lines"] == 1]
    wrap = [r["R"] for r in rows if r["com_lines"] == 2]
    lo = max((w for w in wrap if not keep or w < min(keep)), default=None)
    hi = min(keep, default=None)
    G = rows[0]["G"]; L = rows[0]["L"]
    avail = 468 - L - (P.COL_PT - P.pbefore(G) - hi if hi else 0)
    print(f"{key}: wrap<= {lo}  keep>= {hi}  -> allowance in ({50.40 - hi if hi else None:.2f}, {50.40 - lo if lo else None:.2f})  "
          f"mono={all(not (w > (hi or 1e9)) for w in wrap)}")
