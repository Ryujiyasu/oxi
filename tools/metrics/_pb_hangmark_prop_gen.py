# -*- coding: utf-8 -*-
"""Which line-end marks does the 2026-10 Word hang, for a proportional and a
fixed-pitch face, at compat 14 / no compat declared / compat 15?

_pb_hang_fonts (2026-10-18, c14_paren text, 6 faces): at compat 14 Arial, Times
New Roman and Calibri hang the final '.' / ',' by its full width (left and
justified); Courier New, Consolas and Lucida Console never do; at compat 15
nothing hangs.  This pins the mark set and the no-compat case.

Each arm is one docx: a filler paragraph «AAAA AAAA ... word<mark>» per page,
right indent swept in 5tw steps.  A mark HANGS when the line keeps it with its
glyph starting inside the right indent and ending past it.

    python tools/metrics/_pb_hangmark_prop_gen.py gen
    python tools/metrics/_pb_hangmark_prop_gen.py word
    python tools/metrics/_pb_hangmark_prop_gen.py compare <renderer.exe>   # per-line glyph counts vs Word
"""
import json, os, re, shutil, subprocess, sys, tempfile, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", os.environ.get("PB_OUT", "_pb_hangmark_prop"))
# PB_TAIL: what follows the mark ("" = the mark ends the paragraph)
TAIL = os.environ.get("PB_TAIL", " NEXT")
# PB_LEAD: a first line (ended by <w:br/>) so the filler + mark is line 2
LEAD = os.environ.get("PB_LEAD", "")
MARKS_ALL = {"period": ".", "comma": ",", "semi": ";", "colon": ":", "bang": "!", "query": "?",
         "rparen": ")", "rbracket": "]", "rdquote": "”", "rsquote": "’", "apos": "'",
         "quot": '"', "hyphen": "-", "letter": "x"}
MARKS = {k: v for k, v in MARKS_ALL.items() if k in os.environ.get("PB_MARKS", ",".join(MARKS_ALL)).split(",")}
FONTS = tuple(os.environ.get("PB_FONTS", "Arial,Courier New").split(","))
COMPAT = tuple(os.environ.get("PB_COMPAT", "14,none,15").split(","))
JC = ("left", "both")
RIGHTS = list(range(2050, 2301, 5))   # «twenty<mark>» reaches the edge near 2180tw
if os.environ.get("PB_RIGHTS"):   # "start,stop,step" in twips (narrower faces need a wider sweep)
    RIGHTS = list(range(*map(int, os.environ["PB_RIGHTS"].split(","))))
TEXT = "The regional team reviewed the latest results for participants aged fifteen to twenty"
TEXT = os.environ.get("PB_TEXT", TEXT)   # e.g. a filler ending in a hyphenated token «… 15-17»
# PB_EA: the run's w:eastAsia face (e.g. "ＭＳ 明朝" or "Times New Roman"); empty = not set
EA = ' w:eastAsia="%s"' % os.environ["PB_EA"] if os.environ.get("PB_EA") else ""
# PB_LANG_EA: the run's <w:lang w:eastAsia=...> (e.g. ja-JP / en-US); empty = not set
LANG = '<w:lang w:eastAsia="%s"/>' % os.environ["PB_LANG_EA"] if os.environ.get("PB_LANG_EA") else ""

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"""
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _load_pages(g, t):
    """the glyph dump's pages; the renderer's PNGs (one per page) go with the temp dir."""
    try:
        return json.load(open(g, encoding="utf-8"))["pages"]
    finally:
        shutil.rmtree(t, ignore_errors=True)


def arm(font, cm, jc, mark):
    return "%s_cm%s_%s_%s" % (font.replace(" ", ""), cm, jc, mark)


def build(font, cm, jc, mark):
    ch = MARKS[mark]
    paras = []
    for i, r in enumerate(RIGHTS):
        pb = "<w:pageBreakBefore/>" if i else ""
        jcx = '<w:jc w:val="both"/>' if jc == "both" else ""
        esc = lambda x: x.replace("&", "&amp;").replace('"', "&quot;")
        # PB_LEAD ends with a hard line break, so the filler + mark is the paragraph's LINE 2
        t = (esc(LEAD) + '</w:t><w:br/><w:t xml:space="preserve">' if LEAD else "") + esc(TEXT + ch + TAIL)
        rfonts = '<w:rFonts w:ascii="%s" w:hAnsi="%s"%s/>' % (font, font, EA)
        paras.append('<w:p><w:pPr>%s%s<w:ind w:right="%d"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
                     '<w:r><w:rPr>%s<w:sz w:val="22"/>%s</w:rPr>'
                     '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (pb, jcx, r, rfonts, LANG, t))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:bottom="720" w:left="1080" w:right="800"/>'
           '</w:sectPr></w:body></w:document>' % (W, "".join(paras)))
    compat = "" if cm == "none" else (
        '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%s"/></w:compat>' % cm)
    settings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s>%s</w:settings>' % (W, compat)
    p = os.path.join(OUT, arm(font, cm, jc, mark) + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        ct, drels = CT, DRELS
        # PB_PARTS_FROM=<docx> PB_PARTS=fontTable,theme,styles: copy those parts of a real
        # document into the arm (additive bisection of a document-level switch)
        src = os.environ.get("PB_PARTS_FROM")
        for k, part in enumerate(filter(None, os.environ.get("PB_PARTS", "").split(","))):
            name = {"theme": "word/theme/theme1.xml"}.get(part, "word/%s.xml" % part)
            rtype = {"fontTable": "fontTable", "theme": "theme", "styles": "styles"}[part]
            ctype = {"fontTable": "wordprocessingml.fontTable+xml", "theme": "theme+xml",
                     "styles": "wordprocessingml.styles+xml"}[part]
            data = zipfile.ZipFile(src).read(name).decode("utf-8")
            # PB_STYLES_SUB="regex=>repl;;...": edit the copied styles.xml
            if part == "styles":
                for rule in filter(None, os.environ.get("PB_STYLES_SUB", "").split(";;")):
                    pat, rep = rule.split("=>")
                    data = re.sub(pat, rep, data, flags=re.S)
            z.writestr(name, data.encode("utf-8"))
            drels = drels.replace("</Relationships>", '<Relationship Id="rIdP%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/%s" Target="%s"/></Relationships>' % (k, rtype, name[5:]))
            prefix = "application/vnd.openxmlformats-officedocument." if part != "theme" else "application/vnd.openxmlformats-officedocument."
            ct = ct.replace("</Types>", '<Override PartName="/%s" ContentType="%s%s"/></Types>' % (name, prefix, ctype))
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", doc); z.writestr("word/settings.xml", settings)


def arms():
    return [(f, c, j, m) for f in FONTS for c in COMPAT for j in JC for m in MARKS]


def gen():
    os.makedirs(OUT, exist_ok=True)
    for a in arms():
        build(*a)
    print("built", len(arms()), "docs x", len(RIGHTS), "paragraphs in", OUT)


def limit(r):
    return 612 - 40 - r / 20


def word_hang(docx):
    """max over the sweep of (glyph end - limit) for a line-end mark whose glyph
    starts inside the limit (None when the mark never ends a line past it)."""
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    if not os.path.exists(pdf):
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
    best = None
    for r, pg in zip(RIGHTS, fitz.open(pdf)):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [c for s in l["spans"] for c in s["chars"] if c["c"].strip()]
                if ch and ch[-1]["bbox"][0] < limit(r) and ch[-1]["bbox"][2] > limit(r) + 1.0:
                    best = max(best or 0, round(ch[-1]["bbox"][2] - limit(r), 2))
    return best


def oxi_hang(docx, exe):
    t = tempfile.mkdtemp()
    g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True, timeout=600)
    best = None
    for r, p in zip(RIGHTS, _load_pages(g, t)):
        rows = {}
        for x in p["glyphs"]:
            if x["char"].strip():
                rows.setdefault(round(x["baseline"], 1), []).append(x)
        for v in rows.values():
            last = max(v, key=lambda c: c["x"])
            # glyph end from the next-start convention is unavailable; use x + advance proxy
            end = last["x"] + last.get("w", last["font_size"] * 0.3)
            if last["x"] < limit(r) and end > limit(r) + 1.0:
                best = max(best or 0, round(end - limit(r), 2))
    return best


def run(mode, exe=None):
    res = {}
    for a in arms():
        docx = os.path.join(OUT, arm(*a) + ".docx")
        res[arm(*a)] = word_hang(docx) if mode == "word" else oxi_hang(docx, exe)
    json.dump(res, open(os.path.join(OUT, "_%s.json" % mode), "w"), indent=0)
    for f in FONTS:
        for c in COMPAT:
            for j in JC:
                hung = [m for m in MARKS if res[arm(f, c, j, m)]]
                print("  %-12s cm%-4s %-4s hangs: %s" % (f, c, j, " ".join(hung) or "-"))


if __name__ == "__main__" and sys.argv[1] != "compare":
    m = sys.argv[1]
    if m == "gen":
        gen()
    else:
        run(m, sys.argv[2] if len(sys.argv) > 2 else None)


def line_lengths_word(pdf):
    import fitz
    out = []
    for pg in fitz.open(pdf):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                n = sum(1 for s in l["spans"] for c in s["chars"] if c["c"].strip())
                if n:
                    rows[round(l["bbox"][1])] = rows.get(round(l["bbox"][1]), 0) + n
        out.append([v for _, v in sorted(rows.items())])
    return out


def line_lengths_oxi(docx, exe):
    t = tempfile.mkdtemp()
    g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True, timeout=600)
    out = []
    for p in _load_pages(g, t):
        rows = {}
        for x in p["glyphs"]:
            if x["char"].strip():
                rows[round(x["baseline"])] = rows.get(round(x["baseline"]), 0) + 1
        out.append([v for _, v in sorted(rows.items())])
    return out


def compare(exe):
    """per arm: paragraphs whose per-line glyph counts differ between Word and Oxi."""
    bad = 0
    for a in arms():
        docx = os.path.join(OUT, arm(*a) + ".docx")
        w, o = line_lengths_word(docx[:-5] + ".pdf"), line_lengths_oxi(docx, exe)
        miss = [r for r, x, y in zip(RIGHTS, w, o) if x != y]
        bad += len(miss)
        if miss:
            print("  %-40s %2d/%d differ, e.g. right=%d word %s oxi %s" % (
                arm(*a), len(miss), len(RIGHTS), miss[0], w[RIGHTS.index(miss[0])], o[RIGHTS.index(miss[0])]))
    print("total differing paragraphs:", bad, "of", len(arms()) * len(RIGHTS))


if __name__ == "__main__" and sys.argv[1] == "compare":
    compare(sys.argv[2])
