"""Page-bottom capacity derivation: controlled margin/pitch/fs sweep.

Uniform SINGLE-LINE paragraphs ("第i条") in a typed docGrid. Word's
per-page line capacity = index of the first paragraph on page 2.
Sweeping the BOTTOM margin pins Word's effective bottom acceptance:
    keep line N  iff  top_eff + (N-1)*pitch + THRESH <= 841.9 - bottom/20
The flip margin between N and N+1 lines gives  C = top_eff + THRESH
directly. Sweeping fs separates THRESH(fs) from top_eff; sweeping the
TOP margin separates top_eff.

Usage:
  python _pb_capacity_gen.py gen      -> writes docs to _pb_capacity/
  python _pb_capacity_gen.py measure  -> Word COM: first para of page 2 per doc
"""
import os, sys, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_capacity")

MINCHO = pg.MINCHO


def para(i, sz):
    r = f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="{sz}"/>'
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>第{i}条</w:t></w:r></w:p>')


SENT = pg.SENT  # the probelac 3-line paragraph text (body_paras)


def mpara(i, sz):
    r = f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="{sz}"/>'
    txt = f"第{i}条　{SENT}"
    return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r></w:p>')


def build(pitch, top, bottom, sz, n=60, kind="b"):
    grid = f'<w:docGrid w:type="lines" w:linePitch="{pitch}"/>'
    mar = (f'<w:pgMar w:top="{top}" w:right="1418" w:bottom="{bottom}" '
           f'w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>')
    if kind == "m":
        body = "".join(mpara(i + 1, sz) for i in range(40))
    else:
        body = "".join(para(i + 1, sz) for i in range(n))
    body += pg.sectpr(mar=mar, grid=grid)
    return pg.doc(body)


CASES = []
# bottom sweep at pitch 312 and 360, fs 21 (10.5pt)
for pitch in (312, 360):
    for bottom in range(1300, 1521, 20):
        CASES.append(("b", pitch, 1418, bottom, 21))
# fs sweep at fixed margins (separates THRESH(fs))
for pitch in (312, 360):
    for sz in (18, 21, 22, 24):
        CASES.append(("f", pitch, 1418, 1418, sz))
# top sweep at fixed bottom (separates top_eff)
for top in range(1300, 1521, 40):
    CASES.append(("t", 360, top, 1418, 21))
# FINE bottom sweeps around the single-line (line0/centered) flips: 0.1pt step
for bottom in range(1382, 1403, 2):
    CASES.append(("b", 312, 1418, bottom, 21))
for bottom in range(1416, 1445, 2):
    CASES.append(("b", 360, 1418, bottom, 21))
# MULTI-LINE (3-line SENT paras): last-line threshold, fine bottom sweep.
# Discriminator = y_p2 (first para START on p2): clean top vs 1-line straddle.
for bottom in range(1398, 1439, 4):
    CASES.append(("m", 312, 1418, bottom, 21))
for bottom in range(1440, 1481, 4):
    CASES.append(("m", 360, 1418, bottom, 21))


def name(c):
    kind, pitch, top, bottom, sz = c
    return f"cap_{kind}_p{pitch}_t{top}_b{bottom}_s{sz}.docx"


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for c in CASES:
        kind, pitch, top, bottom, sz = c
        path = os.path.join(OUTDIR, name(c))
        pg.write_docx(path, build(pitch, top, bottom, sz, kind=kind))
    print(f"wrote {len(CASES)} docs to {OUTDIR}")


def measure():
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for c in CASES:
            path = os.path.abspath(os.path.join(OUTDIR, name(c)))
            doc = word.Documents.Open(path, ReadOnly=True)
            try:
                # first paragraph whose (collapsed-start) page == 2
                cap = None
                y2 = None
                n = doc.Paragraphs.Count
                for i in range(1, n + 1):
                    rng = doc.Paragraphs(i).Range
                    st = doc.Range(rng.Start, rng.Start)
                    pgno = st.Information(3)  # wdActiveEndPageNumber on collapsed
                    if pgno >= 2:
                        cap = i - 1
                        y2 = st.Information(6)
                        break
                # y of first + last para on page 1
                y1 = doc.Range(doc.Paragraphs(1).Range.Start,
                               doc.Paragraphs(1).Range.Start).Information(6)
                ylast = None
                if cap:
                    r = doc.Paragraphs(cap).Range
                    ylast = doc.Range(r.Start, r.Start).Information(6)
                rec = {"case": name(c), "kind": c[0], "pitch": c[1], "top": c[2],
                       "bottom": c[3], "sz": c[4], "cap": cap,
                       "y_first": y1, "y_last": ylast, "y_p2": y2}
                results.append(rec)
                print(f"{name(c)}: cap={cap} y1={y1:.2f} ylast={ylast} ")
            finally:
                doc.Close(False)
    finally:
        word.Quit()
    out = os.path.join(OUTDIR, "_results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=1)
    print(f"-> {out}")


if __name__ == "__main__":
    if sys.argv[1:] == ["gen"]:
        gen()
    elif sys.argv[1:] == ["measure"]:
        measure()
    else:
        print(__doc__)
