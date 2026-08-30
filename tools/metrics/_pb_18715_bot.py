# -*- coding: utf-8 -*-
"""CONFOUNDED -- kept for the record; use `_pb_18715_gap.py` instead.

The idea was to read Word's page-5 body limit off `reports__0018715b4769984f`
by sweeping `w:pgMar w:bottom`, which moves the content bottom 1:1. It does --
but it also changes EVERY page's capacity, so the upstream breaks move, and
they move differently in the two engines: at bottom 69pt Word's page 5 starts
26.5pt further into the document than Oxi's, so the two are no longer laying
out the same content and the flip margins are not comparable.

The fix is to move the paragraph without touching the limit: `_pb_18715_gap.py`
sweeps the space AFTER the preceding paragraph (and, with GAP_SELF=1, the
paragraph's own), which leaves the limit, the note set and every upstream break
alone.

    python tools/metrics/_pb_18715_bot.py gen [--sweep lo hi step]
    python tools/metrics/_pb_18715_bot.py word|oxi
"""Measure Word's page-5 body limit on `reports__0018715b4769984f` directly.

Every model so far leaves 0.79pt unexplained at that boundary, so stop modelling
the reserve and read the limit off the document itself: sweep `w:pgMar w:bottom`,
which moves the content bottom 1:1, and find the margin at which Word starts
keeping the paragraph it currently pushes ("However, Professor Phippen says").
The flip margin IS Word's limit, in the same coordinates Oxi reports.

Both engines are read the same way, so the difference between their flip margins
is the whole defect, with no reserve arithmetic in between.

    python tools/metrics/_pb_18715_bot.py gen [--sweep lo hi step]
    python tools/metrics/_pb_18715_bot.py word
    python tools/metrics/_pb_18715_bot.py oxi
"""
import os, re, sys, json, zipfile, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = "pipeline_data/docx_corpus/en/reports/0018715b4769984f.docx"
OUT = r"C:\tmp\pb_18715_bot"
REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
MARK = "However, Professor Phippen"
NOTEMARK = "theguardian.com/education/2025/jan/19"


def parse_sweep(argv):
    if "--sweep" in argv:
        i = argv.index("--sweep")
        return list(range(int(argv[i + 1]), int(argv[i + 2]) + 1, int(argv[i + 3])))
    return list(range(1340, 1445, 4))


def gen(sweep):
    os.makedirs(OUT, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    doc = next(d for it, d in items if it.filename == "word/document.xml").decode("utf-8")
    assert 'w:bottom="1440"' in doc
    for tw in sweep:
        path = os.path.join(OUT, "b%05d.docx" % tw)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for it, data in items:
                if it.filename == "word/document.xml":
                    data = doc.replace('w:bottom="1440"',
                                       'w:bottom="%d"' % tw).encode("utf-8")
                z.writestr(it, data)
    print("built %d docs in %s" % (len(sweep), OUT))


_WORD = [None]


def _word_app():
    """One Word instance for the whole sweep: this document takes ~40s to
    export, and restarting Word per arm both trebled that and left the run
    wedged with no WINWORD process alive."""
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
    page = None
    kept = 0
    first = last = float("nan")
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
        if idx is not None and page is None:
            page = pno + 1
            # the paragraph runs to the end of this page or to its own last line
            kept = len(body) - idx
            first, last = body[idx][0], body[-1][0]
    doc.close()
    return page, kept, first, last


def oxi_read(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    page = None
    kept = 0
    first = last = float("nan")
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
            if not t.strip():
                continue
            rows.setdefault(round(e["y"], 3), []).append((e.get("x", 0), t))
        body, nnotes = [], 0
        for y in sorted(rows):
            txt = "".join(t for _, t in sorted(rows[y]))
            if sep is not None and y > sep:
                nnotes += 1
            else:
                body.append((y, txt))
        idx = next((i for i, (y, t) in enumerate(body)
                    if re.sub(r"\s+", "", t).startswith(re.sub(r"\s+", "", MARK))), None)
        if idx is not None and page is None:
            page = pno + 1
            kept = len(body) - idx
            first, last = body[idx][0], body[-1][0]
    return page, kept, first, last


mode = sys.argv[1] if len(sys.argv) > 1 else "gen"
sw = parse_sweep(sys.argv)
if mode == "gen":
    gen(sw)
else:
    reader = word_read if mode == "word" else oxi_read
    print("%s   bottom-margin sweep: which page holds %r\n" % (mode.upper(), MARK))
    print("  bottom_tw  bottom_pt   para_page  lines_kept   first_y    page_last_y")
    for tw in sw:
        docx = os.path.join(OUT, "b%05d.docx" % tw)
        if not os.path.exists(docx):
            print("  %9d  MISSING" % tw)
            continue
        page, kept, first, last = reader(docx)
        print("  %9d %10.2f %10s %11d %9.2f %12.2f"
              % (tw, tw / 20.0, page, kept, first, last), flush=True)

if mode == "word" and _WORD[0] is not None:
    try:
        _WORD[0].Quit()
    except Exception:
        pass
