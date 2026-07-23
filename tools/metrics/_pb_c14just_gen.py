#!/usr/bin/env python3
"""compat14 Latin justify badness probe (REPORT_compat14_latin_justify_badness §9).

Pins the alt_slack compress/wrap threshold T (bounded to 7.12..14.30pt by the
real-doc Courier quantization) with a 1tw right-indent sweep, and runs the §9.3
exception matrix (short word / sentence boundary / tail). Faithful compat14 host
(Courier New 12, LEGAL, jc=both); wpJustification/usePrinterMetrics ABSENT (both
inert). Word PDF glyph origins are the truth.

    python _pb_c14just_gen.py gen      # build the sweep docx
    python _pb_c14just_gen.py measure   # Word COM lines + PDF glyph analysis
"""
import os, sys, copy, zipfile, json
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
HOST = os.path.join(HERE, "..", "..", "pipeline_data", "docx_corpus", "en",
                    "legal", "0011dcc222423b30.docx")
OUTDIR = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_c14just")
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WNS}
def qn(t): return "{%s}%s" % (WNS, t)

ADV = 7.2012          # Courier New 12pt advance (measured, PDF glyph origin)
HALF_EM = ADV / 2.0   # 3.6006 space floor
COL_PT = 468.0        # LEGAL column: 12240-1440-1440 = 9360tw = 468pt
LEFT_PT = 72.0        # left margin

# Build a jc=both line "wwww wwww ... wwww CAND" with N filler words of `flen`
# chars + a candidate word. Sweep right-indent so alt_slack R = A - P_before
# crosses the flip. P_before = width of the fillers (no trailing space);
# candidate is appended after a single ASCII space.
def build_line(nfill, flen, cand):
    fillers = " ".join("w" * flen for _ in range(nfill))
    return fillers + " " + cand

STEP_TW = 2   # 0.1pt R resolution (1tw feasible but 3x the pages)

def specimens():
    """(name, text, right_indent_tw, meta). Fixed 12 fillers of 4 chars (p_before
    = 424.9pt, fits the 468pt column); vary the candidate to move D/C; sweep the
    right-indent so alt_slack R = A - p_before goes 17 -> 5. If the flip R is the
    same across candidates, T is D/C-independent (local alt-slack threshold)."""
    out = []
    NFILL, FLEN = 12, 4
    p_before_chars = NFILL * FLEN + (NFILL - 1)   # 59 chars
    p_before_pt = p_before_chars * ADV            # 424.87pt
    for lbl, cand in [("and", "and"), ("again", "again"), ("matter", "matter")]:
        text = build_line(NFILL, FLEN, cand)
        cand_pt = len(cand) * ADV
        p_nat_pt = p_before_pt + ADV + cand_pt
        s_count = text.count(" ")
        cap_pt = s_count * (ADV - HALF_EM)
        r_tw = 17 * 20
        idx = 0
        while r_tw >= 5 * 20:
            ri_tw = int(round((COL_PT - p_before_pt) * 20)) - r_tw
            if ri_tw < 0:
                r_tw -= STEP_TW; continue
            a_pt = COL_PT - ri_tw / 20.0
            d_pt = p_nat_pt - a_pt
            out.append((f"{lbl}_R{idx:03d}", text, ri_tw,
                        {"R": round(a_pt - p_before_pt, 3), "D": round(d_pt, 3),
                         "C": round(cap_pt, 3), "DC": round(d_pt / cap_pt, 3) if cap_pt else 0,
                         "S": s_count, "p_before": round(p_before_pt, 2), "cand": cand}))
            idx += 1
            r_tw -= STEP_TW
    return out

def rewrite_settings(raw):
    root = etree.fromstring(raw)
    compat = root.find(".//w:compat", NS)
    for l in ("wpJustification", "usePrinterMetrics"):
        n = compat.find("w:" + l, NS)
        if n is not None:
            compat.remove(n)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

def add_para(body, name, text, ri_tw, bid):
    p = etree.SubElement(body, qn("p"))
    ppr = etree.SubElement(p, qn("pPr"))
    etree.SubElement(ppr, qn("pageBreakBefore"))
    sp = etree.SubElement(ppr, qn("spacing")); sp.set(qn("line"), "480"); sp.set(qn("lineRule"), "auto")
    if ri_tw:
        ind = etree.SubElement(ppr, qn("ind")); ind.set(qn("right"), str(ri_tw))
    etree.SubElement(ppr, qn("jc")).set(qn("val"), "both")
    bs = etree.SubElement(p, qn("bookmarkStart")); bs.set(qn("id"), str(bid)); bs.set(qn("name"), name)
    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t")); t.set("{%s}space" % XML, "preserve"); t.text = text
    etree.SubElement(p, qn("bookmarkEnd")).set(qn("id"), str(bid))

def rewrite_document(raw, specs):
    root = etree.fromstring(raw)
    body = root.find("w:body", NS)
    sect = copy.deepcopy(body.find("w:sectPr", NS))
    for c in list(body):
        body.remove(c)
    for bid, (name, text, ri, meta) in enumerate(specs, 1):
        add_para(body, name, text, ri, bid)
    body.append(sect)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    specs = specimens()
    with zipfile.ZipFile(HOST) as z:
        orig = {n: z.read(n) for n in z.namelist()}
    parts = dict(orig)
    parts["word/settings.xml"] = rewrite_settings(orig["word/settings.xml"])
    parts["word/document.xml"] = rewrite_document(orig["word/document.xml"], specs)
    docx = os.path.abspath(os.path.join(OUTDIR, "sweep.docx"))
    with zipfile.ZipFile(docx, "w", zipfile.ZIP_DEFLATED) as z:
        for pn, data in parts.items():
            z.writestr(pn, data)
    json.dump({n: m for n, _, _, m in specs}, open(os.path.join(OUTDIR, "_meta.json"), "w"), indent=1)
    print("wrote", docx, "with", len(specs), "specimens")

def measure():
    import win32com.client, fitz
    specs = specimens()
    meta = {n: m for n, _, _, m in specs}
    docx = os.path.abspath(os.path.join(OUTDIR, "sweep.docx"))
    pdf = os.path.abspath(os.path.join(OUTDIR, "sweep.pdf"))
    word = win32com.client.DispatchEx("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
    d = word.Documents.Open(docx, ReadOnly=True); d.Repaginate()
    lines = {}
    for name, _, _, _ in specs:
        lines[name] = int(d.Bookmarks(name).Range.ComputeStatistics(1))
    d.ExportAsFixedFormat(pdf, 17); d.Close(False); word.Quit()
    # PDF: per page min space-gap on 1-line pages
    doc = fitz.open(pdf); sg = {}
    order = [n for n, _, _, _ in specs]
    for pi, pg in enumerate(doc):
        if pi >= len(order): break
        chars = []
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for ch in s["chars"]:
                        x, y = ch["origin"]
                        if y < 90: chars.append((round(x, 3), ch["c"]))
        chars.sort()
        gaps = [chars[i+1][0]-chars[i][0] for i in range(len(chars)-1) if chars[i][1] in (' ', ' ')]
        sg[order[pi]] = round(min(gaps), 3) if gaps else None
    res = {n: {"lines": lines[n], "min_space": sg.get(n), **meta[n]} for n in order}
    json.dump(res, open(os.path.join(OUTDIR, "_result.json"), "w"), indent=1)
    # find the flip (1->2 lines) per level, report R at flip
    print("=== flip (1->2 lines) per candidate (D/C level) ===")
    for lvl in ("and", "again", "matter"):
        rows = [(n, res[n]) for n in order if n.startswith(lvl)]
        rows.sort(key=lambda kv: -kv[1]["R"])   # descending R (wide->narrow)
        prev = None
        for n, r in rows:
            if prev is not None and prev["lines"] == 1 and r["lines"] == 2:
                print(f"  {lvl}: FLIP at R in ({r['R']:.3f}, {prev['R']:.3f}]  "
                      f"D/C={r['DC']:.2f} S={r['S']}  (last-fit min_space={prev['min_space']})")
                break
            prev = r
        else:
            counts = {}
            for _, r in rows: counts[r["lines"]] = counts.get(r["lines"], 0)+1
            print(f"  {lvl}: no clean flip in range; line-count dist {counts}, R range "
                  f"[{rows[-1][1]['R']:.2f},{rows[0][1]['R']:.2f}]")

if __name__ == "__main__":
    {"gen": gen, "measure": measure}[sys.argv[1] if len(sys.argv) > 1 else "gen"]()
