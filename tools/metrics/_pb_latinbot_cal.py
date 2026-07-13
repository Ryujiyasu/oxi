"""Calibri variant of _pb_latinbot_gen: distinguishes the FULL-hhea box
page-bottom rule (S827) from the baseline+typo_desc rule discovered at the
framework fn-area boundary. TNR is degenerate (hhea_desc ~= typo_desc);
Calibri separates them by (550-421)/2048*11 = 0.693pt.

Line 49 keep predictions (Letter, top=1440, Calibri 11 hhea 13.4277):
  box model  : keep iff bottom <= 1240.8tw  (flip 1240->1242)
  typo model : keep iff bottom <= 1254.6tw  (flip 1254->1256)

Usage: python _pb_latinbot_cal.py gen | measure
"""
import os, sys, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_latinbot")

CAL = "Calibri"


def para(i):
    r = f'<w:rFonts w:ascii="{CAL}" w:hAnsi="{CAL}"/><w:sz w:val="22"/>'
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">Section {i} text line.</w:t></w:r></w:p>')


def build(bottom, n=70):
    pgsz = '<w:pgSz w:w="12240" w:h="15840"/>'
    mar = (f'<w:pgMar w:top="1440" w:right="1440" w:bottom="{bottom}" '
           f'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>')
    body = "".join(para(i + 1) for i in range(n))
    body += pg.sectpr(pgsz=pgsz, mar=mar, grid='')
    return pg.doc(body)


CASES = list(range(1230, 1263, 2))


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for bottom in CASES:
        p = os.path.join(OUTDIR, f"lbc_b{bottom}.docx")
        pg.write_docx(p, build(bottom), font=CAL, sz="22", compat="15",
                      cpunct=False)
    print("generated", len(CASES))


def measure():
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    out = {}
    try:
        for bottom in CASES:
            nm = f"lbc_b{bottom}"
            p = os.path.abspath(os.path.join(OUTDIR, nm + ".docx"))
            doc = word.Documents.Open(p, ReadOnly=True)
            try:
                first_p2 = None
                for i in range(1, doc.Paragraphs.Count + 1):
                    rng = doc.Paragraphs(i).Range
                    r0 = doc.Range(rng.Start, rng.Start)
                    if r0.Information(3) >= 2:
                        first_p2 = i
                        break
                out[nm] = first_p2
                print(nm, "->", first_p2, flush=True)
            finally:
                doc.Close(False)
    finally:
        word.Quit()
    with open(os.path.join(OUTDIR, "_measure_cal.json"), "w") as f:
        json.dump(out, f, indent=1)


if __name__ == "__main__":
    if sys.argv[1:] == ["gen"]:
        gen()
    elif sys.argv[1:] == ["measure"]:
        measure()
    else:
        print(__doc__)
