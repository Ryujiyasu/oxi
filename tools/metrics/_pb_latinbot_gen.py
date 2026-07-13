"""No-grid LATIN page-bottom threshold derivation (the nyserda LM0 wall).

Uniform SINGLE-LINE Latin paragraphs ("Section i text.") in a NO-docGrid
Letter page, Times New Roman 12pt, single spacing, compat-15 settings.
Word's per-page line capacity = index of the first paragraph on page 2.
Line i top = top_margin + (i-1) * pitch, pitch = hhea 1.1499 x 12 = 13.799.
Sweeping the BOTTOM margin in 2tw (0.1pt) steps pins Word's acceptance:
    keep line N  iff  top_N + THRESH <= page_h - bottom/20
The S775 nyserda reading suggested THRESH ~= baseline advance (13.8) from
the NEXT line's baseline test, vs Oxi's top + height <= bottom. The fs
sweep separates THRESH(fs).

Usage:
  python _pb_latinbot_gen.py gen      -> pipeline_data/_pb_latinbot/
  python _pb_latinbot_gen.py measure  -> Word COM: first para of page 2
"""
import os, sys, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_latinbot")

TNR = "Times New Roman"


def para(i, sz):
    r = f'<w:rFonts w:ascii="{TNR}" w:hAnsi="{TNR}"/><w:sz w:val="{sz}"/>'
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">Section {i} text line.</w:t></w:r></w:p>')


def build(top, bottom, sz, n=70):
    # Letter page like nyserda (12240x15840), NO docGrid.
    pgsz = '<w:pgSz w:w="12240" w:h="15840"/>'
    mar = (f'<w:pgMar w:top="{top}" w:right="1440" w:bottom="{bottom}" '
           f'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>')
    body = "".join(para(i + 1, sz) for i in range(n))
    body += pg.sectpr(pgsz=pgsz, mar=mar, grid='')
    return pg.doc(body)


CASES = []
# coarse bottom sweep at fs 24 (12pt): 1300..1520 step 20tw (1pt)
for bottom in range(1300, 1521, 20):
    CASES.append((1440, bottom, 24))
# fs sweep at fixed margins
for sz in (20, 22, 28):
    for bottom in range(1360, 1481, 40):
        CASES.append((1440, bottom, sz))


def name(top, bottom, sz):
    return f"lb_t{top}_b{bottom}_s{sz}"


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for top, bottom, sz in CASES:
        p = os.path.join(OUTDIR, name(top, bottom, sz) + ".docx")
        pg.write_docx(p, build(top, bottom, sz), font=TNR, sz=str(24),
                      compat="15", cpunct=False)
    print("generated", len(CASES))


def measure():
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    out = {}
    try:
        for top, bottom, sz in CASES:
            nm = name(top, bottom, sz)
            p = os.path.abspath(os.path.join(OUTDIR, nm + ".docx"))
            doc = word.Documents.Open(p, ReadOnly=True)
            try:
                first_p2 = None
                for i in range(1, doc.Paragraphs.Count + 1):
                    rng = doc.Paragraphs(i).Range
                    r0 = doc.Range(rng.Start, rng.Start)
                    pgno = r0.Information(3)  # wdActiveEndPageNumber via collapsed start
                    if pgno >= 2:
                        first_p2 = i
                        break
                out[nm] = first_p2
                print(nm, "->", first_p2, flush=True)
            finally:
                doc.Close(False)
    finally:
        word.Quit()
    with open(os.path.join(OUTDIR, "_measure.json"), "w") as f:
        json.dump(out, f, indent=1)


if __name__ == "__main__":
    if sys.argv[1:] == ["gen"]:
        gen()
    elif sys.argv[1:] == ["measure"]:
        measure()
    else:
        print(__doc__)
