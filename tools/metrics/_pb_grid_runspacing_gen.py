# -*- coding: utf-8 -*-
"""How does a run's w:spacing move characters on a linesAndChars grid?

ja/reference__0ea3ec86 (docGrid linesAndChars linePitch 411 charSpace 2048, ＭＳ 明朝 11pt
body) has «　聴覚障害者等…» with w:spacing -4: Word steps 11.52 per character (= the grid
cell, as with no spacing at all), Oxi 10.46.  A coarse sweep on the faithful slice read
Word -40:7.668 -20:9.600 -10:10.967 -4:11.520 -1:11.508 0:11.508 +4:12.120 +10:12.792
+20:13.510 -- not linear in the spacing.

This builds the slice once (the paragraph, the document's styles / settings / section) and
repeats the paragraph with every spacing in PB_SPACINGS, one per page; per arm Word's mean
advance over ten ideographs.

    python tools/metrics/_pb_grid_runspacing_gen.py
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import feature_census as FC  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_grid_runspacing")
DID = os.environ.get("PB_DOC", "ja/reference__0ea3ec86480140c2")
NEEDLE = os.environ.get("PB_NEEDLE", "聴覚障害者等との交流")
SPACINGS = [int(v) for v in os.environ.get("PB_SPACINGS", ",".join(str(v) for v in range(-40, 41, 2))).split(",")]
SZ = os.environ.get("PB_SZ")   # half-points: also pin the run size
# PB_NOBAL=1: drop <w:balanceSingleByteDoubleByteWidth/> (0ea3ec86 carries it)
NOBAL = os.environ.get("PB_NOBAL") == "1"


def build():
    z = zipfile.ZipFile(FC.load()[DID]["path"])
    parts = {n: z.read(n) for n in z.namelist()}
    d = parts["word/document.xml"].decode("utf-8")
    i = d.find(NEEDLE)
    s = max(d.rfind("<w:p ", 0, i), d.rfind("<w:p>", 0, i)); e = d.find("</w:p>", i) + 6
    p = d[s:e]
    j = d.find("<w:sectPr", i)
    sect = re.search(r"<w:sectPr[ >].*?</w:sectPr>", d[j:], re.S).group(0)
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    paras = []
    for k, v in enumerate(SPACINGS):
        q = re.sub(r'<w:spacing w:val="-?\d+"/>', '<w:spacing w:val="%d"/>' % v, p)
        if SZ:
            q = q.replace("<w:rPr>", '<w:rPr><w:sz w:val="%s"/>' % SZ)
        if k:
            q = q.replace("<w:pPr>", "<w:pPr><w:pageBreakBefore/>", 1)
        paras.append(q)
    if NOBAL:
        parts["word/settings.xml"] = parts["word/settings.xml"].replace(b"<w:balanceSingleByteDoubleByteWidth/>", b"")
    b0 = re.search(r"<w:body[^>]*>", d).end()
    parts["word/document.xml"] = (d[:b0] + "".join(paras) + sect + "</w:body></w:document>").encode("utf-8")
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "grs%s%s.docx" % ("_sz" + SZ if SZ else "", "_nobal" if NOBAL else ""))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return path


def advances(pdf):
    import fitz
    out = []
    for pg in fitz.open(pdf):
        best = None
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                cs = [c for s in l["spans"] for c in s["chars"] if c["c"].strip()]
                if len(cs) >= 11:
                    best = cs
                    break
            if best:
                break
        out.append((best[10]["origin"][0] - best[0]["origin"][0]) / 10 if best else None)
    return out


if __name__ == "__main__":
    import win32com.client
    docx = build(); pdf = docx[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    oxi = None
    if os.environ.get("PB_OXI"):
        import subprocess, tempfile, json
        t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
        subprocess.run([os.environ["PB_OXI"], docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
        oxi = []
        for pg in json.load(open(g, encoding="utf-8"))["pages"]:
            gl = [c for c in pg["glyphs"] if c["char"].strip()]
            y0 = min(c["baseline"] for c in gl)
            L = sorted([c for c in gl if abs(c["baseline"] - y0) < 1], key=lambda c: c["x"])
            oxi.append((L[10]["x"] - L[0]["x"]) / 10 if len(L) > 10 else None)
    for k, (v, a) in enumerate(zip(SPACINGS, advances(pdf))):
        o = oxi[k] if oxi and k < len(oxi) else None
        print("spacing %4d tw (%+.2fpt)  Word advance %s   Oxi %s" % (v, v / 20, "%.3f" % a if a else None, "%.3f" % o if o else None), flush=True)
