#!/usr/bin/env python3
"""compat14 Latin justify — downstream-context isolation probe
(REPORT_compat14_latin_justify_downstream.md §6).

S995's residual: 4 NEW terminal errors where the SAME local tuple (after-capacity
excess ~3.68 or ~7.28pt) gives Word WRAP in one paragraph and KEEP in another
(Pair A: 581 WRAP / 660 KEEP; Pair B: twin147 WRAP / 105 KEEP). The candidate-local
geometry is identical, so the discriminator is DOWNSTREAM/paragraph context.

This probe FIXES the terminal line (12 "mmmm" body words + a candidate, over the
half-em capacity via a right-indent that puts it below the single-line flip R~10.7
so Word WRAPs it standalone) and varies ONE downstream axis at a time, with N
"aaaa" filler lines (13 words/line, verified) preceding it so the terminal always
starts fresh on line N+1 with an IDENTICAL local tuple:

  1. preceding full-line count : N = 0,1,2,3,5,8   (line ordinal / paragraph length)
  2. page proximity            : terminal at page top / mid / last-2-lines / after-break
  3. tail cardinality          : 0 / 1 / 2 words after the candidate (is "terminal" the axis?)
  4. paragraph split A/B        : same string as 1 para vs 2 paras (para-global vs page-global?)

Faithful compat14 host (Courier New 12, LEGAL, jc=both), wpJustification /
usePrinterMetrics ABSENT (both inert). Word PDF glyph origins are the truth.

    python _pb_c14down_gen.py gen       # build the probe docx
    python _pb_c14down_gen.py measure    # Word COM lines + PDF glyph analysis
"""
import os, sys, copy, zipfile, json
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
HOST = os.path.join(HERE, "..", "..", "pipeline_data", "docx_corpus", "en",
                    "legal", "0011dcc222423b30.docx")
OUTDIR = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_c14down")
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WNS}
def qn(t): return "{%s}%s" % (WNS, t)

ADV = 7.2012          # Courier New 12pt advance (measured, PDF glyph origin)
HALF_EM = ADV / 2.0
COL_PT = 468.0        # LEGAL column: 12240-1440-1440 = 9360tw = 468pt
NFILL, FLEN = 12, 4   # terminal body: 12 "mmmm" words (matches _pb_c14just p_before)
P_BEFORE = NFILL * FLEN * ADV + (NFILL - 1) * ADV   # 424.87pt (no candidate)

# 13 "aaaa" words = 64 chars = 460.9pt < COL, 14th wraps -> exactly 1 line each.
FILLER_WORD = "aaaa"
FILLER_PER_LINE = 13

def terminal(cand):
    return " ".join("mmmm" for _ in range(NFILL)) + " " + cand

def para_of_lines(n):
    """n full filler lines (n*13 'aaaa' words)."""
    return " ".join(FILLER_WORD for _ in range(n * FILLER_PER_LINE))

def ri_for_R(R):
    """right-indent (tw) so the terminal's room R = available - P_BEFORE."""
    return int(round((COL_PT - P_BEFORE - R) * 20))

# CLEAN construction: paragraph = "mmmm" × (12·(N+1)) + candidate. avail = 424.87+R
# holds exactly 12 "mmmm" per line (stable for R∈[2,12]), so lines 1..N are full and
# the candidate's room R on the LAST mmmm-line (line N+1) is identical across N. The
# only thing that varies is the line ORDINAL of the terminal. Sweep R low→high per N
# to find the WRAP/KEEP flip per N (single-line flip = 10.73 from _pb_c14just).
CAND_W = {"or": 2, "final.": 6, "partnership.": 12}
R_SWEEP = [2, 4, 6, 7, 8, 9, 9.5, 10, 10.25, 10.5, 10.75, 11, 12]
N_SWEEP = [0, 1, 2, 4]

def clean_para(cand, n):
    """n+1 mmmm-lines (12 each) then the candidate on the last mmmm-line (room R)."""
    return " ".join("mmmm" for _ in range(NFILL * (n + 1))) + " " + cand

def specimens():
    out = []
    # --- Core: WRAP/KEEP flip vs R, per terminal line-ordinal (N preceding lines) ---
    for n in N_SWEEP:
        for R in R_SWEEP:
            ri = ri_for_R(R)
            out.append((f"E_N{n}_R{int(round(R*100)):04d}", [(clean_para("final.", n), "both", None)],
                        {"axis": "ordinal", "N": n, "R": R, "cand": "final.", "ri": ri}))
    # --- Candidate-width: does the flip R depend on candidate width? (N=2) ---
    for cand in ("or", "partnership."):
        for R in [2, 6, 9, 10.5, 12, 20, 30, 45]:
            ri = ri_for_R(R)
            out.append((f"W_{cand.rstrip('.')}_R{int(round(R*100)):04d}", [(clean_para(cand, 2), "both", None)],
                        {"axis": "cand_width", "N": 2, "R": R, "cand": cand, "ri": ri}))
    # --- Axis 4: split A/B at R=9 (single WRAPs, multi KEEPs) ---
    ri = ri_for_R(9.0)
    line1 = " ".join("mmmm" for _ in range(NFILL))
    out.append((f"D_1para", [(clean_para("final.", 1), "both", None)],
                {"axis": "split", "R": 9.0, "form": "1para", "cand": "final.", "ri": ri}))
    out.append((f"D_2para", [(line1, "both", None), (terminal("final."), "both", None)],
                {"axis": "split", "R": 9.0, "form": "2para", "cand": "final.", "ri": ri}))
    return out

def rewrite_settings(raw):
    root = etree.fromstring(raw)
    compat = root.find(".//w:compat", NS)
    for l in ("wpJustification", "usePrinterMetrics"):
        n = compat.find("w:" + l, NS)
        if n is not None:
            compat.remove(n)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

def add_para(body, name, text, jc, first_line, ri_tw, bid, first_of_group):
    p = etree.SubElement(body, qn("p"))
    ppr = etree.SubElement(p, qn("pPr"))
    if first_of_group:
        etree.SubElement(ppr, qn("pageBreakBefore"))
    sp = etree.SubElement(ppr, qn("spacing")); sp.set(qn("line"), "480"); sp.set(qn("lineRule"), "auto")
    if ri_tw:
        ind = etree.SubElement(ppr, qn("ind")); ind.set(qn("right"), str(ri_tw))
    if first_line is not None:
        etree.SubElement(ppr, qn("ind")).set(qn("firstLine"), str(first_line))
    etree.SubElement(ppr, qn("jc")).set(qn("val"), jc)
    if name is not None:
        bs = etree.SubElement(p, qn("bookmarkStart")); bs.set(qn("id"), str(bid)); bs.set(qn("name"), name)
    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t")); t.set("{%s}space" % XML, "preserve"); t.text = text
    if name is not None:
        etree.SubElement(p, qn("bookmarkEnd")).set(qn("id"), str(bid))

def rewrite_document(raw, specs):
    root = etree.fromstring(raw)
    body = root.find("w:body", NS)
    sect = copy.deepcopy(body.find("w:sectPr", NS))
    for c in list(body):
        body.remove(c)
    bid = 0
    for name, paras, meta in specs:
        ri = meta["ri"]
        for i, (text, jc, fl) in enumerate(paras):
            bid += 1
            # only the FIRST para of each specimen carries the bookmark (measured)
            add_para(body, name if i == 0 else None, text, jc, fl, ri, bid, i == 0)
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
    docx = os.path.abspath(os.path.join(OUTDIR, "down.docx"))
    with zipfile.ZipFile(docx, "w", zipfile.ZIP_DEFLATED) as z:
        for pn, data in parts.items():
            z.writestr(pn, data)
    json.dump({n: m for n, _, m in specs}, open(os.path.join(OUTDIR, "_meta.json"), "w"), indent=1)
    print("wrote", docx, "with", len(specs), "specimens")
    print("R_WRAP ri=%d  R_KEEP ri=%d  P_BEFORE=%.2f" % (ri_for_R(R_WRAP), ri_for_R(R_KEEP), P_BEFORE))

def measure():
    import win32com.client, fitz
    specs = specimens()
    meta = {n: m for n, _, m in specs}
    docx = os.path.abspath(os.path.join(OUTDIR, "down.docx"))
    pdf = os.path.abspath(os.path.join(OUTDIR, "down.pdf"))
    word = win32com.client.DispatchEx("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
    d = word.Documents.Open(docx, ReadOnly=True); d.Repaginate()
    com = {}
    for name, _, _ in specs:
        try:
            com[name] = int(d.Bookmarks(name).Range.ComputeStatistics(1))  # wdStatisticLines
        except Exception as ex:
            com[name] = "ERR:" + str(ex)
    d.ExportAsFixedFormat(pdf, 17); d.Close(False); word.Quit()
    # PDF: per specimen, count the body lines of the FIRST paragraph and read the
    # terminal line (does the candidate 'final.' sit ALONE on the last line = WRAP,
    # or at the end of a full content line = KEEP?). Each specimen's first para
    # starts on its own page (pageBreakBefore).
    doc = fitz.open(pdf)
    order = [n for n, _, _ in specs]
    # collect lines per page (body only)
    pages = []
    for pg in doc:
        lines = {}
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    y = round(s["origin"][1], 1)
                    if y > 760: continue
                    e = lines.setdefault(y, [9e9, -9e9, ""])
                    e[0] = min(e[0], s["bbox"][0]); e[1] = max(e[1], s["bbox"][2]); e[2] += s["text"]
        pages.append([(y, round(x0, 1), round(x1, 1), t) for y, (x0, x1, t) in sorted(lines.items()) if t.strip()])
    res = {}
    for i, name in enumerate(order):
        m = meta[name]
        if m["axis"] == "split" and m["form"] == "2para":
            # two paragraphs share the specimen; the terminal is the 2nd para.
            pg = pages[i] if i < len(pages) else []
        else:
            pg = pages[i] if i < len(pages) else []
        # last line of the (first) paragraph's page: is the candidate alone (WRAP)?
        last = pg[-1] if pg else None
        cand = m.get("cand", "final.")
        candalone = bool(last and last[3].strip() == cand)
        # min inter-word gap on the terminal line (compression signal), PDF glyph origin
        res[name] = {"com_lines": com[name], "pdf_body_lines": len(pg),
                     "last_line": last[3][:40] if last else None,
                     "last_x1": last[2] if last else None,
                     "cand_alone": candalone, **m}
    json.dump(res, open(os.path.join(OUTDIR, "_result.json"), "w"), indent=1)
    # report — flip R per terminal line-ordinal N
    print("=== WRAP(W)/KEEP(K) vs terminal room R, per line-ordinal N (Word) ===")
    print("    single-line (N=0) flip = 10.73; a LOWER flip at N>=1 = line-ordinal shifts it")
    for n in N_SWEEP:
        row = []
        for R in R_SWEEP:
            r = res.get(f"E_N{n}_R{int(round(R*100)):04d}")
            if r: row.append(f"{R:>5.2f}:{'W' if r['cand_alone'] else 'K'}")
        print(f"  N={n}: " + "  ".join(row))
    print("=== Candidate-width: flip R per candidate (N=2) ===")
    for cand in ("or", "partnership."):
        row = []
        for R in [2, 6, 9, 10.5, 12, 20, 30, 45]:
            r = res.get(f"W_{cand.rstrip('.')}_R{int(round(R*100)):04d}")
            if r: row.append(f"{R:>4.1f}:{'W' if r['cand_alone'] else 'K'}")
        print(f"  {cand:>13}: " + "  ".join(row))
    print("=== Axis 4: split A/B (R=9: single WRAPs, multi KEEPs) ===")
    for f in ("1para", "2para"):
        r = res.get(f"D_{f}")
        if r: print(f"  {f}: {'WRAP' if r['cand_alone'] else 'KEEP'}  com_lines={r['com_lines']}  last='{r['last_line']}'")

if __name__ == "__main__":
    {"gen": gen, "measure": measure}[sys.argv[1] if len(sys.argv) > 1 else "gen"]()
