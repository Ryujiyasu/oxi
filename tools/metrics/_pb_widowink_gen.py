"""No-type docGrid LATIN single-spacing 2-line-paragraph widow page-bottom
derivation (the reference__00525b22 wall).

reference__00525b22's references are Calibri 12pt SINGLE-spacing 2-line
paragraphs (date+title wraps) in a no-TYPE docGrid (linePitch=360) Letter
page. Word FITS a 2-line entry whose last-line INK bottom (baseline +
glyph-descent) fits the content bottom; Oxi's widow look-ahead (S608) uses
the last line's NATURAL box (ascent+descent) → the ~0.5pt-taller box
overflows content_bottom by a fraction and widowControl pushes the whole
entry → a +1 cascade through the references.

This probe fills a page with 2-line Calibri-12 paragraphs and sweeps the
BOTTOM margin in 2tw (0.1pt) steps. Word's per-page 2-line-para capacity =
the index of the first paragraph on page 2. The flip point pins the last
line's page-bottom threshold: at the flip,
    keep the 2-line para iff  last_line_top + THRESH <= page_h - bottom/20
THRESH = natural(asc+desc) (Oxi/S608) vs ink(asc+glyph_desc) (hypothesis).

Usage:
  python _pb_widowink_gen.py gen      -> pipeline_data/_pb_widowink/
  python _pb_widowink_gen.py measure  -> Word COM: first para of page 2 (per case)
  python _pb_widowink_gen.py oxi      -> Oxi: first para of page 2 (per case)
  python _pb_widowink_gen.py read     -> flip points Word vs Oxi
"""
import os, sys, json, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_widowink")

FONT = "Calibri"

# A long paragraph that wraps to exactly 2 lines at this width, tagged with a
# unique index so the page-2-first-para can be identified.
def para2(i):
    r = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="24"/>'
    txt = (f"P{i:03d} start of a two line reference entry that wraps to the "
           f"second line here for the widow page bottom probe test now.")
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r></w:p>')


def build(bottom, n=40):
    # Letter page, no-TYPE docGrid linePitch=360, margins like reference__00525b22.
    body = "".join(para2(i + 1) for i in range(n))
    sect = (f'<w:sectPr>'
            f'<w:pgSz w:w="12240" w:h="15840"/>'
            f'<w:pgMar w:top="1152" w:right="1296" w:bottom="{bottom}" '
            f'w:left="1296" w:header="720" w:footer="720" w:gutter="0"/>'
            f'<w:docGrid w:linePitch="360"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{body}{sect}</w:body></w:document>')
    ct = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
    rels = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    styles = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    return doc, ct, rels, drels, styles


# bottom margin sweep: 2tw (0.1pt) steps over ~2pt window around 1152 (the doc's value).
CASES = list(range(1100, 1221, 2))  # 1100..1220tw = 55.0..61.0pt


def name(bottom):
    return f"wi_b{bottom}"


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for bottom in CASES:
        doc, ct, rels, drels, styles = build(bottom)
        p = os.path.join(OUTDIR, f"{name(bottom)}.docx")
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", rels)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles)
    print(f"wrote {len(CASES)} docs to {OUTDIR}")


def measure():
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for bottom in CASES:
            src = os.path.abspath(os.path.join(OUTDIR, f"{name(bottom)}.docx"))
            d = word.Documents.Open(src, ReadOnly=True)
            # first paragraph whose start is on page 2
            first_p2 = None
            for p in d.Paragraphs:
                rng = p.Range
                pg = d.Range(rng.Start, rng.Start).Information(3)  # wdActiveEndPageNumber
                if pg >= 2:
                    t = rng.Text.strip()[:5]
                    first_p2 = t
                    break
            res[bottom] = first_p2
            d.Close(False)
            print(f"  b{bottom}: p2 starts {first_p2}", flush=True)
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, "_word.json"), "w"))
    print("wrote _word.json")


def oxi():
    import subprocess, glob
    RND = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",
          "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe"))
    res = {}
    for bottom in CASES:
        src = os.path.join(OUTDIR, f"{name(bottom)}.docx")
        dump = os.path.join(OUTDIR, f"{name(bottom)}.json")
        subprocess.run([RND, src, os.path.join(OUTDIR, "o_"), f"--dump-layout={dump}"],
                       capture_output=True)
        d = json.load(open(dump, encoding="utf-8"))
        # first "Pnnn" text on page 2 (0-indexed page 1)
        first = None
        if len(d["pages"]) >= 2:
            for e in sorted(d["pages"][1]["elements"], key=lambda e: (e.get("y", 0), e.get("x", 0))):
                t = (e.get("text") or "").strip()
                if t.startswith("P") and t[1:4].isdigit():
                    first = t[:5]
                    break
        res[bottom] = first
        print(f"  b{bottom}: p2 starts {first}", flush=True)
    json.dump(res, open(os.path.join(OUTDIR, "_oxi.json"), "w"))
    print("wrote _oxi.json")


def read():
    w = json.load(open(os.path.join(OUTDIR, "_word.json")))
    o = json.load(open(os.path.join(OUTDIR, "_oxi.json"))) if os.path.exists(os.path.join(OUTDIR, "_oxi.json")) else {}
    print(f"{'bottom':>7} {'pt':>6} {'Word':>6} {'Oxi':>6}")
    for bottom in CASES:
        wv = w.get(str(bottom)) or w.get(bottom)
        ov = o.get(str(bottom)) or o.get(bottom)
        print(f"{bottom:>7} {bottom/20:>6.1f} {str(wv):>6} {str(ov):>6}")
    print("\nThe flip (p2-first index increments) pins each engine's page-bottom")
    print("acceptance. Word flipping at a SMALLER bottom margin than Oxi = Word")
    print("fits a lower last line = Word uses a SHORTER (ink) last-line box.")


if __name__ == "__main__":
    {"gen": gen, "measure": measure, "oxi": oxi, "read": read}[sys.argv[1]]()
