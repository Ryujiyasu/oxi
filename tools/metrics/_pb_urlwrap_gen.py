"""Consecutive-URL line-break derivation (the reference__0014acda wall).

reference__0014acda's "Belgium - FWO" entry holds TWO space-separated
hyperlink URLs. Word render-truth (PDF): it FILLS line 1 with URL1's start
(breaking at a hyphen, "all-|calls") but wraps the WHOLE second URL to a
fresh line, leaving line 2 short (x=374.9 of a 519 margin) = 4 lines.
Oxi's S783 (Latin hyphen = word boundary) shatters URL2 at its hyphens, so
the pieces greedily fill line 2 = 3 lines -> the following para stays at the
page-1 bottom where Word pushes it to page 2 (the -1).

The blanket fix (S862: URL hyphens -> break opportunity for ALL URLs)
OVERSHOOTS -- it wraps URL1 whole too (Belgium 5 lines vs Word 4) and
regressed uk_framework 42->43 pages. So the rule is NOT "URLs never break at
hyphens". Two candidate rules remain:

  H1 (cohesion) : a URL that FOLLOWS another URL wraps WHOLE to a fresh
                  line, regardless of the room left.
  H2 (threshold): ANY long token wraps whole when the remaining room is
                  below some amount (Belgium's URL2 had only ~144pt).

DECISIVE SWEEP: hold URL2 fixed (long, hyphenated) and sweep URL1's length
so the room left on line 1 varies from ~40pt to ~380pt. Then read, from the
Word PDF, whether URL2's first glyph lands on line 1 (FILL) or line 2 (WRAP).
  H1 predicts WRAP at every room value.
  H2 predicts FILL once the room exceeds the threshold.
Control cases pin the other axes:
  ctl_single : label + ONE url            -> expect FILL (S783 is right)
  ctl_word   : label + url + long hyphen-compound (NOT a url)
               -> is the rule URL-specific, or "any long token after a URL"?
  ctl_wordurl: label + long hyphen-compound + url
               -> is it "preceded by a URL", or "url wraps when room small"?

Usage:
  python _pb_urlwrap_gen.py gen      -> pipeline_data/_pb_urlwrap/*.docx
  python _pb_urlwrap_gen.py measure  -> Word PDF (fitz): per-line x/text
"""
import os, sys, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_urlwrap")

FONT = "Calibri"
SZ = "22"  # 11pt, matching reference__0014acda's docDefaults

# URL2: long + hyphenated; longer than a full line so it must break either way.
URL2 = ("https://www.example.org/support-programmes/all-calls/"
        "postdoctoral-researchers/senior-postdoctoral-fellowship/")


def url1(k):
    """A hyphenated URL whose length grows with k (k path segments)."""
    segs = "/".join(["aaaa-bbbb"] * k)
    return f"https://www.example.org/{segs}/" if k else "https://www.example.org/"


def mkpara(text):
    r = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="{SZ}"/>'
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr>'
            f'<w:t xml:space="preserve">{pg.esc(text)}</w:t></w:r></w:p>')


def build(text):
    # A4 + 72pt margins + NO docGrid == reference__0014acda's geometry.
    pgsz = '<w:pgSz w:w="11906" w:h="16838"/>'
    mar = ('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
           'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>')
    return pg.doc(mkpara(text) + pg.sectpr(pgsz=pgsz, mar=mar, grid=''))


LONGWORD = "alpha-bravo-charlie-delta-echo-foxtrot-golf-hotel-india-juliett-kilo"

CASES = {}
# the sweep: URL1 length k -> room left before URL2
for k in range(0, 8):
    CASES[f"sw_k{k}"] = f"L: {url1(k)} {URL2}"
# controls
CASES["ctl_single"] = f"L: {url1(3)}"
CASES["ctl_word"] = f"L: {url1(3)} {LONGWORD}-{LONGWORD}"
CASES["ctl_wordurl"] = f"L: {LONGWORD}-{LONGWORD} {URL2}"


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for nm, text in CASES.items():
        p = os.path.join(OUTDIR, nm + ".docx")
        pg.write_docx(p, build(text), font=FONT, sz=SZ, compat="15", cpunct=False)
    print("generated", len(CASES), "->", os.path.abspath(OUTDIR))


def measure():
    import win32com.client
    import fitz
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    out = {}
    try:
        for nm in CASES:
            src = os.path.abspath(os.path.join(OUTDIR, nm + ".docx"))
            pdf = os.path.abspath(os.path.join(OUTDIR, nm + ".pdf"))
            doc = word.Documents.Open(src, ReadOnly=True)
            try:
                doc.ExportAsFixedFormat(pdf, 17)
            finally:
                doc.Close(False)
            d = fitz.open(pdf)
            lines = []
            for blk in d[0].get_text("dict")["blocks"]:
                for ln in blk.get("lines", []):
                    t = "".join(s["text"] for s in ln["spans"])
                    if t.strip():
                        lines.append((round(ln["bbox"][1], 1), round(ln["bbox"][0], 1),
                                      round(ln["bbox"][2], 1), t))
            d.close()
            lines.sort()
            out[nm] = lines
            print(f"\n== {nm}")
            for y, x0, x1, t in lines:
                print(f"   y={y:6.1f} x=[{x0:5.1f},{x1:5.1f}] {t[:72]!r}")
    finally:
        word.Quit()
    with open(os.path.join(OUTDIR, "_measure.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, indent=1)
    print("\nwrote _measure.json")


if __name__ == "__main__":
    if sys.argv[1:] == ["gen"]:
        gen()
    elif sys.argv[1:] == ["measure"]:
        measure()
    else:
        print(__doc__)
