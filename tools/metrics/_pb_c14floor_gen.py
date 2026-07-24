#!/usr/bin/env python3
"""compat14 justify terminal KEEP/WRAP — pin the per-GAP compression floor C
(REPORT_compat14_latin_justify_downstream.md §8.5, the pre-implement gate).

_pb_c14down established: the discriminator is overflow/compressible_gaps with a
floor C < half-em, ~3.0pt/gap. This probe pins three things the §8.5 gate needs:

  1. C vs gap-count : single-line "mmmm×G + final.", sweep R -> flip R*(G);
     per_gap = (space+cand - R*)/G. Is C constant across G=6/9/12?
  2. double-space weight : G=9 with D literal double-spaces ("  "), sweep R.
     If a "  " has the SAME capacity as a single gap, the flip shifts by 0;
     if 2x, by C per double. Yields the weight W (gap = single + W*double).
  3. period hang : candidate "final." (period, hangable) vs "finalx" (no punct)
     at G=9. Does the hang add ~1 char (7.2pt) of allowance?

Single-line paragraph => no earlier-line repacking; the line is the last (ragged)
line, which compat14 STILL compresses to keep the candidate (the R~10.73 flip in
_pb_c14down was exactly this). Faithful Courier-12 compat14 host, both flags off.

    python _pb_c14floor_gen.py gen
    python _pb_c14floor_gen.py measure
"""
import os, sys, copy, zipfile, json
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
HOST = os.path.join(HERE, "..", "..", "pipeline_data", "docx_corpus", "en",
                    "legal", "0011dcc222423b30.docx")
OUTDIR = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_c14floor")
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WNS}
def qn(t): return "{%s}%s" % (WNS, t)

ADV = 7.2012
COL_PT = 468.0

def pbefore(G):
    """natural width of G 'mmmm' words (no candidate): G*4 chars + (G-1) spaces."""
    return (G * 4 + (G - 1)) * ADV

def ri_for_R(G, R):
    return int(round((COL_PT - pbefore(G) - R) * 20))

def line_text(G, cand, ndouble=0):
    """G 'mmmm' words then candidate; the first ndouble inter-word gaps are doubled."""
    words = ["mmmm"] * G
    parts = [words[0]]
    for i in range(1, G):
        parts.append("  " if i <= ndouble else " ")
        parts.append(words[i])
    parts.append(" " + cand)
    return "".join(parts)

def specimens():
    out = []
    # 1. C vs gap-count: single-line "mmmm×G final.", sweep R across the flip.
    for G in (6, 9, 12):
        flip = 50.4 - 3.3 * G   # expected flip R (C~3.3)
        for R in [flip + d for d in (-3, -2, -1, -0.5, 0, 0.5, 1, 2, 3)]:
            if R <= 0 and COL_PT - pbefore(G) - R > 468: continue
            ri = ri_for_R(G, R)
            if ri < 0: continue
            out.append((f"G{G}_R{int(round(R*100)):+05d}", line_text(G, "final."),
                        {"axis": "gap_count", "G": G, "R": round(R, 2), "cand": "final.", "ndouble": 0, "ri": ri}))
    # 2. double-space weight: G=9, D doubles. Each double adds 7.2 to the line body,
    # so the nominal-R flip rises. weight-1 -> flip ~ 50.4 + D*7.2 - (9+D)*3.5
    # (26 for D=2, 34 for D=4); weight-2 -> flip ~ 19 for both. Sweep wide.
    G = 9
    for D in (0, 2, 4):
        base = 50.4 + D * 7.2 - (9 + D) * 3.5   # weight-1 expectation
        for R in [base + d for d in (-4, -2, -1, 0, 1, 2, 4, 8, 12)]:
            ri = ri_for_R(G, R)
            if ri < 0: continue
            out.append((f"D{D}_R{int(round(R*100)):+06d}", line_text(G, "final.", D),
                        {"axis": "double", "G": G, "D": D, "R": round(R, 2), "cand": "final.", "ndouble": D, "ri": ri}))
    # 3. period hang: G=9, 'final.' (period) vs 'finalx' (no punct). flip ~18.
    G = 9
    for cand in ("final.", "finalx"):
        for R in [12 + d for d in range(0, 12)]:   # 12..23
            ri = ri_for_R(G, R)
            if ri < 0: continue
            out.append((f"H_{cand}_R{int(round(R*100)):+06d}", line_text(G, cand),
                        {"axis": "hang", "G": G, "R": round(float(R), 2), "cand": cand, "ndouble": 0, "ri": ri}))
    # 4. candidate width: G=9, cand width 6/9/12 chars, sweep R -> flip per-gap C(width)?
    # (real 'partnership.' 12ch wraps at per-gap 3.2 < the 6ch controlled 3.58.)
    G = 9
    for cw, cand in ((6, "final."), (9, "finalxyz."), (12, "finalxyzabc.")):
        need = ADV + cw * ADV   # space + candidate
        flip = need - 9 * 3.5   # weight-independent expectation at C~3.5
        for R in [flip + d for d in (-4, -3, -2, -1, 0, 1, 2, 3, 4)]:
            ri = ri_for_R(G, R)
            if ri < 0: continue
            out.append((f"CW{cw}_R{int(round(R*100)):+06d}", line_text(G, cand),
                        {"axis": "cwidth", "G": G, "cw": cw, "need": round(need, 2), "R": round(R, 2), "cand": cand, "ndouble": 0, "ri": ri}))
    return out

def rewrite_settings(raw):
    root = etree.fromstring(raw); compat = root.find(".//w:compat", NS)
    for l in ("wpJustification", "usePrinterMetrics"):
        n = compat.find("w:" + l, NS)
        if n is not None: compat.remove(n)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

def add_para(body, name, text, ri_tw, bid):
    p = etree.SubElement(body, qn("p")); ppr = etree.SubElement(p, qn("pPr"))
    etree.SubElement(ppr, qn("pageBreakBefore"))
    sp = etree.SubElement(ppr, qn("spacing")); sp.set(qn("line"), "480"); sp.set(qn("lineRule"), "auto")
    if ri_tw:
        ind = etree.SubElement(ppr, qn("ind")); ind.set(qn("right"), str(ri_tw))
    etree.SubElement(ppr, qn("jc")).set(qn("val"), "both")
    bs = etree.SubElement(p, qn("bookmarkStart")); bs.set(qn("id"), str(bid)); bs.set(qn("name"), name)
    r = etree.SubElement(p, qn("r")); t = etree.SubElement(r, qn("t"))
    t.set("{%s}space" % XML, "preserve"); t.text = text
    etree.SubElement(p, qn("bookmarkEnd")).set(qn("id"), str(bid))

def rewrite_document(raw, specs):
    root = etree.fromstring(raw); body = root.find("w:body", NS)
    sect = copy.deepcopy(body.find("w:sectPr", NS))
    for c in list(body): body.remove(c)
    for bid, (name, text, meta) in enumerate(specs, 1):
        add_para(body, name, text, meta["ri"], bid)
    body.append(sect)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

def gen():
    os.makedirs(OUTDIR, exist_ok=True); specs = specimens()
    with zipfile.ZipFile(HOST) as z: orig = {n: z.read(n) for n in z.namelist()}
    parts = dict(orig)
    parts["word/settings.xml"] = rewrite_settings(orig["word/settings.xml"])
    parts["word/document.xml"] = rewrite_document(orig["word/document.xml"], specs)
    docx = os.path.abspath(os.path.join(OUTDIR, "floor.docx"))
    with zipfile.ZipFile(docx, "w", zipfile.ZIP_DEFLATED) as z:
        for pn, data in parts.items(): z.writestr(pn, data)
    json.dump({n: m for n, _, m in specs}, open(os.path.join(OUTDIR, "_meta.json"), "w"), indent=1)
    print("wrote", docx, "with", len(specs), "specimens")

def measure():
    import win32com.client, fitz
    specs = specimens(); meta = {n: m for n, _, m in specs}
    docx = os.path.abspath(os.path.join(OUTDIR, "floor.docx"))
    pdf = os.path.abspath(os.path.join(OUTDIR, "floor.pdf"))
    word = win32com.client.DispatchEx("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
    d = word.Documents.Open(docx, ReadOnly=True); d.Repaginate()
    com = {n: int(d.Bookmarks(n).Range.ComputeStatistics(1)) for n, _, _ in specs}
    d.ExportAsFixedFormat(pdf, 17); d.Close(False); word.Quit()
    doc = fitz.open(pdf); order = [n for n, _, _ in specs]; pages = []
    for pg in doc:
        lines = {}
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    y = round(s["origin"][1], 1)
                    if y > 760: continue
                    e = lines.setdefault(y, [9e9, -9e9, ""]); e[0] = min(e[0], s["bbox"][0]); e[1] = max(e[1], s["bbox"][2]); e[2] += s["text"]
        pages.append([(round(x0, 1), round(x1, 1), t) for y, (x0, x1, t) in sorted(lines.items()) if t.strip()])
    res = {}
    for i, name in enumerate(order):
        pg = pages[i] if i < len(pages) else []
        last = pg[-1] if pg else None
        cand = meta[name]["cand"]
        wrap = bool(last and last[2].strip() == cand)   # candidate alone = WRAP
        res[name] = {"com_lines": com[name], "wrap": wrap, "last": last[2][:24] if last else None, **meta[name]}
    json.dump(res, open(os.path.join(OUTDIR, "_result.json"), "w"), indent=1)
    def flip_R(names):
        rows = sorted([res[n] for n in names], key=lambda r: r["R"])
        prev = None
        for r in rows:
            if prev is not None and prev["wrap"] and not r["wrap"]:
                return (prev["R"], r["R"])
            prev = r
        return None
    print("=== 1. C vs gap-count (candidate 'final.', space+cand=50.40) ===")
    for G in (6, 9, 12):
        names = [n for n in order if meta[n]["axis"] == "gap_count" and meta[n]["G"] == G]
        f = flip_R(names)
        if f:
            mid = (f[0] + f[1]) / 2; C = (50.40 - mid) / G
            print(f"  G={G:>2}: flip R in ({f[0]:.2f},{f[1]:.2f}] -> per_gap C = {C:.3f}pt")
        else:
            print(f"  G={G:>2}: no flip; " + " ".join(f"{res[n]['R']:.1f}:{'W' if res[n]['wrap'] else 'K'}" for n in names))
    print("=== 2. double-space weight (G=9, +D doubles; extra width = D*7.2) ===")
    for D in (0, 2, 4):
        names = [n for n in order if meta[n]["axis"] == "double" and meta[n]["D"] == D]
        f = flip_R(names)
        # with D doubles the line's natural is +D*7.2 vs single-only; overflow at room R is
        # (50.40 + D*7.2 - R); gaps if double counts as 1 = 9; capacity units = 9 + W*D
        if f:
            mid = (f[0] + f[1]) / 2; ov = 50.40 + D * 7.2 - mid
            print(f"  D={D}: flip R in ({f[0]:.2f},{f[1]:.2f}] -> overflow={ov:.2f}pt (units 9 + W*{D})")
        else:
            print(f"  D={D}: no flip; " + " ".join(f"{res[n]['R']:.1f}:{'W' if res[n]['wrap'] else 'K'}" for n in names))
    print("=== 3. period hang (G=9): 'final.' vs 'finalx' ===")
    for cand in ("final.", "finalx"):
        names = [n for n in order if meta[n]["axis"] == "hang" and meta[n]["cand"] == cand]
        f = flip_R(names)
        if f: print(f"  {cand:>7}: flip R in ({f[0]:.2f},{f[1]:.2f}]")
        else: print(f"  {cand:>7}: no flip; " + " ".join(f"{res[n]['R']:.1f}:{'W' if res[n]['wrap'] else 'K'}" for n in names))
    print("=== 4. candidate width (G=9, 9 gaps): does the per-gap floor drop with width? ===")
    for cw in (6, 9, 12):
        names = [n for n in order if meta[n]["axis"] == "cwidth" and meta[n]["cw"] == cw]
        f = flip_R(names)
        if f:
            need = meta[names[0]]["need"]; mid = (f[0] + f[1]) / 2; C = (need - mid) / 9
            print(f"  cw={cw:>2} (need={need:.1f}): flip R in ({f[0]:.2f},{f[1]:.2f}] -> per_gap C = {C:.3f}pt")
        else:
            print(f"  cw={cw:>2}: no flip; " + " ".join(f"{res[n]['R']:.1f}:{'W' if res[n]['wrap'] else 'K'}" for n in names))

if __name__ == "__main__":
    {"gen": gen, "measure": measure}[sys.argv[1] if len(sys.argv) > 1 else "gen"]()
