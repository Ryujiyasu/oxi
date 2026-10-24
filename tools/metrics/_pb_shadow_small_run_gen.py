# -*- coding: utf-8 -*-
"""How much does a w14:shadow on a SMALL run grow a line whose height a bigger run sets?

ja/reference__5593c6a7's title box: a 40pt run and a 20pt run with blur 4 / dist 3 (down)
shadow on the 20pt run. Word's line is the plain 40pt line (the shadow stays inside the
big run's box); Oxi adds the 20pt run's shadow extents to the line (+2.8pt).
OXI_TEXT_SHADOW_LINE (`_pb_shadow_line_gen.py`) was derived with the shadow on the run
that sets the line.

Faithful slice of the document's own text box (its bodyPr / insets / style a3 kept):
P1 = «宮城県柴田» at BIG pt + «普通科» at SMALL pt (shadow on the small run), P2 = 22pt.
Sweeps SMALL at BIG 40 for the document's shadow, and the shadow (blur, dist) at SMALL 30.
Prints the P1->P2 baseline pitch minus the no-shadow pitch (the growth), Word and Oxi.

    python tools/metrics/_pb_shadow_small_run_gen.py        (PB_OXI=<renderer.exe> adds Oxi)
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import feature_census as FC  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_shadow_small_run")
FONT = os.environ.get("PB_FONT", "ＭＳ ゴシック")


def slice_parts():
    src = [v["path"] for k, v in FC.load().items() if "5593c6a7" in k][0]
    zi = zipfile.ZipFile(src)
    parts = {n: zi.read(n) for n in zi.namelist()}
    x = parts["word/document.xml"].decode()
    i = x.find("宮城県柴田高等学校")
    a = x.rfind("<mc:AlternateContent>", 0, i)
    b = x.find("</mc:AlternateContent>", i) + len("</mc:AlternateContent>")
    anchor = re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", x[a:b], flags=re.S)
    return parts, x, anchor


def shadow_xml(blur_pt, dist_pt):
    return ('<w14:shadow w14:blurRad="%d" w14:dist="%d" w14:dir="5400000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="t">'
            '<w14:srgbClr w14:val="000000"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:shadow>' % (round(blur_pt * 12700), round(dist_pt * 12700)))


def para(runs, jc="both", style=None):
    ppr = "<w:pPr>" + ('<w:pStyle w:val="%s"/>' % style if style else "") + '<w:jc w:val="%s"/></w:pPr>' % jc
    body = "".join('<w:r><w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hint="eastAsia"/><w:sz w:val="%d"/>%s</w:rPr><w:t>%s</w:t></w:r>'
                   % (FONT, FONT, int(pt * 2), extra, text) for pt, text, extra in runs)
    return "<w:p>" + ppr + body + "</w:p>"


def build(name, big, small, shadow):
    parts, x, anchor = slice_parts()
    tc0 = anchor.find("<w:txbxContent>") + len("<w:txbxContent>")
    tc1 = anchor.find("</w:txbxContent>")
    content = para([(big, "宮城県柴田", ""), (small, "普通科", shadow)]) + para([(22, "柴高生の志と地域貢献", "")], "center")
    anc = anchor[:tc0] + content + anchor[tc1:]
    body0 = x.find("<w:body>") + 8
    sect = x.rfind("<w:sectPr")
    nx = x[:body0] + "<w:p><w:r>" + anc + "</w:r></w:p>" + x[sect:]
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, name + ".docx")
    parts = dict(parts)
    parts["word/document.xml"] = nx.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, d in parts.items():
            zo.writestr(n, d)
    return path


def pitch_word(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    ys = {}
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = "".join(c["c"] for c in s["chars"]).strip()
                if t:
                    ys.setdefault(t[:4], s["chars"][0]["origin"][1])
    return ys["柴高生の"] - ys["宮城県柴"]


def pitch_oxi(exe, docx):
    import subprocess, json, tempfile
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    rows = {}
    for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"]:
        rows.setdefault(round(c["baseline"], 1), []).append(c)
    ys = {}
    for k in sorted(rows):
        t = "".join(c["char"] for c in sorted(rows[k], key=lambda c: c["x"])).strip()
        if t:
            ys.setdefault(t[:4], k)
    return ys["柴高生の"] - ys["宮城県柴"]


if __name__ == "__main__":
    import win32com.client
    arms = [("none", 40, 20, "")]
    for small in (10, 16, 20, 24, 30, 36, 40):
        arms.append(("s%d_b4d3" % small, 40, small, shadow_xml(4, 3)))
    for blur, dist in ((8, 6), (0, 6), (8, 0), (2, 1.5)):
        arms.append(("s30_b%gd%g" % (blur, dist), 40, 30, shadow_xml(blur, dist)))
    base_w = base_o = None
    for name, big, small, shadow in arms:
        p = build(name, big, small, shadow); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        pw = pitch_word(pdf)
        po = pitch_oxi(os.environ["PB_OXI"], p) if os.environ.get("PB_OXI") else None
        if name == "none":
            base_w, base_o = pw, po
        print("%-12s big %d small %2d  Word pitch %6.2f growth %+5.2f | Oxi %s" % (
            name, big, small, pw, pw - base_w,
            ("pitch %6.2f growth %+5.2f" % (po, po - base_o)) if po is not None else "-"), flush=True)
