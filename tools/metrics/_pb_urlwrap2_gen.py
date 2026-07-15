"""Consecutive-URL wrap: pin the ROOM threshold (stage 2 of _pb_urlwrap_gen).

Stage 1 falsified the "cohesion" hypothesis (a URL following a URL does NOT
always wrap whole -- sw_k0..k3/k7 FILL). The behaviour tracks the ROOM left on
the line before the 2nd URL:
    k3      room 162.3 -> FILL
    Belgium room 148.4 -> WRAP   (reference__0014acda, real doc)
    k4      room 110.5 -> WRAP
so the flip is in (148.4, 162.3]. It is NOT "the first break-piece fits":
Belgium's first '/'-piece is ~130pt and fits in 148.4, yet Word wrapped.

This stage sweeps the room in fine steps with a FILLER of n 'x' chars, and
pairs every case with a CONTROL doc (same text minus URL2) so the room is
MEASURED (control line-1 x1) rather than estimated:
    room = content_right - x1(control)
Two URL2 variants (different first-segment length) test whether the threshold
is URL-dependent or a property of the line; two page widths (W1 default A4,
W2 narrower right margin) test absolute-pt vs proportional-to-line-width.

Usage:
  python _pb_urlwrap2_gen.py gen
  python _pb_urlwrap2_gen.py measure   -> prints room vs FILL/WRAP per case
"""
import os, sys, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_urlwrap2")

FONT = "Calibri"
SZ = "22"  # 11pt

# Long hyphenated URLs (both far longer than one line -> must break either way).
URL2_A = ("https://www.example.org/support-programmes/all-calls/"
          "postdoctoral-researchers/senior-postdoctoral-fellowship/")
# variant with a SHORTER leading '/'-piece (short domain, like fwo.be)
URL2_B = ("https://www.fwo.be/en/support-programmes/all-calls/"
          "postdoctoral-researchers/senior-postdoctoral-fellowship/")

URL1 = "https://www.example.org/aaaa-bbbb/"

# right margin (tw): W1 = default 1440 (content right 523.3pt),
# W2 = 2880 (content right 451.3pt) -> tests proportional vs absolute
WIDTHS = {"W1": 1440, "W2": 2880}


def mkpara(text):
    r = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="{SZ}"/>'
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr>'
            f'<w:t xml:space="preserve">{pg.esc(text)}</w:t></w:r></w:p>')


def build(text, right):
    pgsz = '<w:pgSz w:w="11906" w:h="16838"/>'
    mar = (f'<w:pgMar w:top="1440" w:right="{right}" w:bottom="1440" '
           f'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>')
    return pg.doc(mkpara(text) + pg.sectpr(pgsz=pgsz, mar=mar, grid=''))


def text_for(n, url2):
    pad = "x" * n
    body = f"{pad} {URL1}" if n else URL1
    return (f"{body} {url2}", body)  # (main, control)


CASES = {}
for wk, right in WIDTHS.items():
    for uk, url2 in (("A", URL2_A), ("B", URL2_B)):
        # n sweep: each 'x' ~5.2pt in Calibri 11 -> ~5pt room steps
        for n in range(0, 26):
            CASES[f"{wk}_{uk}_n{n:02d}"] = (n, url2, right)


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for nm, (n, url2, right) in CASES.items():
        main, ctl = text_for(n, url2)
        pg.write_docx(os.path.join(OUTDIR, nm + ".docx"), build(main, right),
                      font=FONT, sz=SZ, compat="15", cpunct=False)
        pg.write_docx(os.path.join(OUTDIR, nm + "_c.docx"), build(ctl, right),
                      font=FONT, sz=SZ, compat="15", cpunct=False)
    print("generated", len(CASES) * 2, "->", os.path.abspath(OUTDIR))


def _lines(word, fitz, path):
    pdf = path[:-5] + ".pdf"
    doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
    try:
        doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)
    finally:
        doc.Close(False)
    d = fitz.open(pdf)
    out = []
    for blk in d[0].get_text("dict")["blocks"]:
        for ln in blk.get("lines", []):
            t = "".join(s["text"] for s in ln["spans"])
            if t.strip():
                out.append((round(ln["bbox"][1], 1), round(ln["bbox"][0], 1),
                            round(ln["bbox"][2], 1), t))
    d.close()
    out.sort()
    return out


def measure():
    import win32com.client, fitz
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for nm, (n, url2, right) in CASES.items():
            content_right = (11906 - right) / 20.0
            ctl = _lines(word, fitz, os.path.join(OUTDIR, nm + "_c.docx"))
            main = _lines(word, fitz, os.path.join(OUTDIR, nm + ".docx"))
            if not ctl or not main:
                continue
            url1_end = ctl[0][2]           # control line-1 x1 == URL1 end
            room = round(content_right - url1_end, 1)
            # FILL iff line 1 of the main doc is longer than the control's
            # line 1 (i.e. some of URL2 landed on line 1)
            fill = main[0][2] > url1_end + 1.0
            res[nm] = dict(n=n, room=room, fill=bool(fill),
                           url1_end=url1_end, line1_x1=main[0][2],
                           nlines=len(main))
            print(f"{nm}: room={room:6.1f}  {'FILL' if fill else 'WRAP'}", flush=True)
    finally:
        word.Quit()
    with open(os.path.join(OUTDIR, "_measure.json"), "w", encoding="utf-8") as f:
        json.dump(res, f, indent=1)
    # summarize flip per (width,url)
    print("\n=== flip points ===")
    for wk in WIDTHS:
        for uk in ("A", "B"):
            rows = sorted((v for k, v in res.items() if k.startswith(f"{wk}_{uk}_")),
                          key=lambda r: r["room"])
            lo = max((r["room"] for r in rows if not r["fill"]), default=None)
            hi = min((r["room"] for r in rows if r["fill"]), default=None)
            cr = (11906 - WIDTHS[wk]) / 20.0
            line_w = cr - 72.0
            print(f"{wk}/{uk}: WRAP<= {lo} < FILL >= {hi}   (line_w={line_w:.1f}, "
                  f"ratio {None if hi is None else round(hi/line_w,3)})")


if __name__ == "__main__":
    if sys.argv[1:] == ["gen"]:
        gen()
    elif sys.argv[1:] == ["measure"]:
        measure()
    else:
        print(__doc__)
