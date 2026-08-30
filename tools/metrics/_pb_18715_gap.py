# -*- coding: utf-8 -*-
"""Move ONE paragraph across page 5's bottom on `reports__0018715b4769984f`.

The bottom-margin sweep was confounded: it changes every page's capacity, so
Word's upstream breaks move and page 5 stops holding the same content as Oxi's.
Shrinking the space AFTER the preceding paragraph instead leaves the limit, the
note set and every upstream break untouched, and slides only the paragraph under
test ("However, Professor Phippen says") up the page in 0.2pt steps.

The document's default is w:after=160 (8pt), so X in [0,160] buys 8pt of travel
upward. The X at which each engine flips from "3 lines on page 5" to "pushed
whole" is its body limit, read in the same coordinates.

    python tools/metrics/_pb_18715_gap.py gen [--sweep lo hi step]
    python tools/metrics/_pb_18715_gap.py word
    python tools/metrics/_pb_18715_gap.py oxi
"""
import os, re, sys, json, zipfile, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = "pipeline_data/docx_corpus/en/reports/0018715b4769984f.docx"
REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
MARK = "However, Professor Phippen"
PREV_TAIL = "information spreading on their platforms"
# GAP_SELF=1 puts the swept w:after on the paragraph UNDER TEST instead of the
# one before it. If Word's page-bottom test counts a paragraph's own trailing
# space, this sweep moves the flip 1:1 while leaving the paragraph's position
# alone; if it does not, the flip never moves. The two sweeps must not share a
# directory: same arm names, different documents, and the PDFs are cached.
SELF = os.environ.get("GAP_SELF") == "1"
OUT = r"C:\tmp\pb_18715_gap" + ("_self" if SELF else "")


def parse_sweep(argv):
    if "--sweep" in argv:
        i = argv.index("--sweep")
        return list(range(int(argv[i + 1]), int(argv[i + 2]) + 1, int(argv[i + 3])))
    return list(range(0, 161, 8))


def patch(doc, after_tw):
    """Set w:after on the paragraph that ends with PREV_TAIL."""
    paras = list(re.finditer(r"<w:p(?: [^>]*)?>.*?</w:p>", doc, re.S))
    hit = None
    for m in paras:
        body = re.sub(r"<[^>]+>", "", m.group(0))
        if (MARK in body) if SELF else (PREV_TAIL in body):
            hit = m
            if SELF:
                break
    assert hit is not None, "target paragraph not found"
    p = hit.group(0)
    spacing = '<w:spacing w:after="%d"/>' % after_tw
    if "<w:pPr>" in p:
        if re.search(r"<w:spacing[^>]*/>", p):
            p2 = re.sub(r"<w:spacing[^>]*/>", spacing, p, count=1)
        else:
            p2 = p.replace("<w:pPr>", "<w:pPr>" + spacing, 1)
    else:
        p2 = re.sub(r"(<w:p(?: [^>]*)?>)", r"\1<w:pPr>" + spacing + "</w:pPr>", p, count=1)
    return doc[:hit.start()] + p2 + doc[hit.end():]


def gen(sweep):
    os.makedirs(OUT, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    doc = next(d for it, d in items
               if it.filename == "word/document.xml").decode("utf-8")
    for tw in sweep:
        patched = patch(doc, tw).encode("utf-8")
        with zipfile.ZipFile(os.path.join(OUT, "a%04d.docx" % tw), "w",
                             zipfile.ZIP_DEFLATED) as z:
            for it, data in items:
                z.writestr(it, patched if it.filename == "word/document.xml" else data)
    print("built %d docs in %s" % (len(sweep), OUT))


_WORD = [None]


def _word_app():
    import win32com.client
    if _WORD[0] is None:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        _WORD[0] = app
    return _WORD[0]


def word_read(docx):
    import fitz
    pdf = docx[:-5] + ".pdf"
    if not os.path.exists(pdf):
        app = _word_app()
        d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        try:
            d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
        finally:
            d.Close(False)
    doc = fitz.open(pdf)
    out = (None, 0, float("nan"), float("nan"))
    for pno in range(doc.page_count):
        pg = doc[pno]
        rules = [dr["rect"].y0 for dr in pg.get_drawings()
                 if dr["rect"].width > 50 and dr["rect"].height < 4 and dr["rect"].y0 > 300]
        rule = min(rules) if rules else None
        rows = []
        for blk in pg.get_text("dict")["blocks"]:
            for l in blk.get("lines", []):
                sp = [s for s in l["spans"] if s["text"].strip()]
                if sp:
                    rows.append((sp[0]["origin"][1],
                                 "".join(s["text"] for s in sp).strip()))
        rows.sort()
        body = [(y, t) for y, t in rows if rule is None or y <= rule]
        idx = next((i for i, (y, t) in enumerate(body) if t.startswith(MARK)), None)
        if idx is not None:
            out = (pno + 1, len(body) - idx, body[idx][0], body[-1][0])
            break
    doc.close()
    return out


def oxi_read(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    for pno, pg in enumerate(d["pages"]):
        sep = None
        for e in pg["elements"]:
            if (e.get("text") or "").strip():
                continue
            if e.get("y", 0) > 300 and 100 < e.get("w", 0) < 200 and e.get("h", 9) < 4:
                sep = e["y"] if sep is None else min(sep, e["y"])
        rows = {}
        for e in pg["elements"]:
            t = (e.get("text") or "")
            if t.strip():
                rows.setdefault(round(e["y"], 3), []).append((e.get("x", 0), t))
        body = [(y, "".join(t for _, t in sorted(rows[y])))
                for y in sorted(rows) if sep is None or y <= sep]
        idx = next((i for i, (y, t) in enumerate(body)
                    if re.sub(r"\s+", "", t).startswith(re.sub(r"\s+", "", MARK))), None)
        if idx is not None:
            return pno + 1, len(body) - idx, body[idx][0], body[-1][0]
    return None, 0, float("nan"), float("nan")


mode = sys.argv[1] if len(sys.argv) > 1 else "gen"
sw = parse_sweep(sys.argv)
if mode == "gen":
    gen(sw)
else:
    reader = word_read if mode == "word" else oxi_read
    print("%s   prev-paragraph w:after sweep (default 160 = 8pt)\n" % mode.upper())
    print("  after_tw   after_pt   page  lines_kept    first_y   page_last_y")
    for tw in sw:
        docx = os.path.join(OUT, "a%04d.docx" % tw)
        if not os.path.exists(docx):
            print("  %8d  MISSING" % tw)
            continue
        page, kept, first, last = reader(docx)
        print("  %8d %10.2f %6s %11d %10.2f %13.2f"
              % (tw, tw / 20.0, page, kept, first, last), flush=True)
    if mode == "word" and _WORD[0] is not None:
        try:
            _WORD[0].Quit()
        except Exception:
            pass
