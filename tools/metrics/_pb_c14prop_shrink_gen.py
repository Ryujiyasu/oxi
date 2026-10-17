# -*- coding: utf-8 -*-
"""Does the 2026-10 Word squeeze a PROPORTIONAL compat-14 justified line?

The Courier probes (_pb_c14floor re-run, 2026-10-18) found no squeeze at all
without wpJustification.  tests/fixtures/list_pagination/c14_paren.docx (Arial
11, jc both, compat 14, no wpJustification) still keeps «85%.» on line 2 with
its last glyph 5.6pt past the right indent.  This sweeps that paragraph's right
indent in 10tw steps and records where «85%.» leaves line 2, for
{Arial, Times New Roman, Courier New} x {compat 14, 15} x {both, left}; the
left-aligned flip is the natural width, so (left flip - both flip) is the squeeze.

    python tools/metrics/_pb_c14prop_shrink_gen.py gen
    python tools/metrics/_pb_c14prop_shrink_gen.py word     # each docx in a fresh Word
    python tools/metrics/_pb_c14prop_shrink_gen.py oxi <renderer.exe>
"""
import json, os, re, shutil, subprocess, sys, tempfile, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
SRC = os.path.join(REPO, "tests", "fixtures", "list_pagination", "c14_paren.docx")
OUT = os.path.join(REPO, "pipeline_data", os.environ.get("PB_OUT", "_pb_c14prop_shrink"))
FONTS = tuple(os.environ.get("PB_FONTS", "Arial,Times New Roman,Courier New").split(","))
COMPAT = (14, 15)
JC = ("both", "left")
RIGHTS = list(range(900, 1501, 10))          # w:ind w:right sweep (twips)


def _load_pages(g, t):
    """the glyph dump's pages; the renderer's PNGs (one per page) go with the temp dir."""
    try:
        return json.load(open(g, encoding="utf-8"))["pages"]
    finally:
        shutil.rmtree(t, ignore_errors=True)


def arm(font, cm, jc):
    return "%s_cm%d_%s" % (font.replace(" ", ""), cm, jc)


def gen():
    os.makedirs(OUT, exist_ok=True)
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    doc = parts["word/document.xml"].decode()
    p0 = doc[doc.find("<w:p>"):doc.find("</w:p>") + 6]
    for font in FONTS:
        for cm in COMPAT:
            for jc in JC:
                paras = []
                for r in RIGHTS:
                    p = p0.replace('w:right="1200"', 'w:right="%d"' % r)
                    p = p.replace('<w:jc w:val="both"/>', '' if jc == "left" else '<w:jc w:val="both"/>')
                    p = p.replace('w:ascii="Arial" w:hAnsi="Arial"', 'w:ascii="%s" w:hAnsi="%s"' % (font, font))
                    paras.append(p)
                body = doc[:doc.find("<w:p>")] + "".join(paras) + doc[doc.find("</w:p>") + 6:]
                out = dict(parts)
                out["word/document.xml"] = body.encode()
                out["word/settings.xml"] = parts["word/settings.xml"].replace(b'w:val="14"', b'w:val="%d"' % cm)
                with zipfile.ZipFile(os.path.join(OUT, arm(font, cm, jc) + ".docx"), "w", zipfile.ZIP_DEFLATED) as w:
                    for n, b in out.items():
                        w.writestr(n, b)
    print("built", len(FONTS) * len(COMPAT) * len(JC), "docs x", len(RIGHTS), "paragraphs in", OUT)


def keeps_from_lines(pages):
    """pages: [[line text, ...], ...] with one paragraph per page (pageBreakBefore).
    True when «85%.» is NOT alone on a line (it stayed on line 2)."""
    res = []
    for lines in pages:
        lines = [l for l in lines if not re.fullmatch(r"[\d.)]+", l.strip())]
        res.append(not any(l.strip().endswith("85%.") and l.strip().startswith("85%") for l in lines))
    return res


def word_pages(docx):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    pages = []
    for pg in fitz.open(pdf):
        rows = {}
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip():
                        rows.setdefault(round(s["origin"][1], 0), []).append((s["origin"][0], s["text"]))
        pages.append(["".join(t for _, t in sorted(v)) for _, v in sorted(rows.items())])
    return pages


def oxi_pages(docx, exe):
    t = tempfile.mkdtemp()
    g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True, timeout=600)
    pages = []
    for p in _load_pages(g, t):
        rows = {}
        for x in p["glyphs"]:
            rows.setdefault(round(x["baseline"], 0), []).append(x)
        pages.append(["".join(c["char"] for c in sorted(v, key=lambda c: c["x"])) for _, v in sorted(rows.items())])
    return pages


def hangs(pdf):
    """per page: the line-end glyphs that START inside the right indent but END past it."""
    import fitz
    out = []
    for r, pg in zip(RIGHTS, fitz.open(pdf)):
        limit = 612 - 40 - r / 20
        hung = []
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [c for s in l["spans"] for c in s["chars"] if c["c"].strip()]
                if ch and ch[-1]["bbox"][2] > limit + 0.05:
                    hung.append((ch[-1]["c"], round(ch[-1]["bbox"][0] - limit, 2), round(ch[-1]["bbox"][2] - limit, 2)))
        out.append(hung)
    return out


def flip(keeps):
    """smallest right indent (twips) at which «85%.» stays on line 2 -> the boundary."""
    k = [r for r, ok in zip(RIGHTS, keeps) if ok]
    return max(k) if k else None


def run(mode, exe=None):
    res = {}
    for font in FONTS:
        for cm in COMPAT:
            for jc in JC:
                a = arm(font, cm, jc)
                docx = os.path.join(OUT, a + ".docx")
                pages = word_pages(docx) if mode == "word" else oxi_pages(docx, exe)
                res[a] = keeps_from_lines(pages)
    json.dump(res, open(os.path.join(OUT, "_%s.json" % mode), "w"), indent=0)
    print("%s: largest right indent (tw) that still keeps «85%%.» on line 2" % mode)
    for font in FONTS:
        for cm in COMPAT:
            b, l = flip(res[arm(font, cm, "both")]), flip(res[arm(font, cm, "left")])
            sq = (b - l) / 20 if b is not None and l is not None else None
            print("  %-16s cm%d  both %s  left %s  -> squeeze %s pt" % (font, cm, b, l, sq))


if __name__ == "__main__":
    m = sys.argv[1]
    if m == "gen":
        gen()
    else:
        run(m, sys.argv[2] if len(sys.argv) > 2 else None)
