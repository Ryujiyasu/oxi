# -*- coding: utf-8 -*-
"""Does an EMPTY paragraph grow its line by the paragraph mark's w14:shadow?

ja/reference__5593c6a7's second text box: «…（野球、陸上…» / an empty paragraph whose
pPr/rPr carries the box's shadow / «他にも有名で誇れることが！». Word sets the last line
2.6pt lower than Oxi, which gives an empty paragraph the plain line of its mark size.

Faithful slice of the document's title box: P1 «部活の充実» 22pt, an empty paragraph
(mark rPr per arm), P3 «他にも有名» 22pt. Prints the P1->P3 baseline pitch for Word and
Oxi; the growth over the plain arm is the mark shadow's contribution.

    python tools/metrics/_pb_mark_shadow_empty_gen.py       (PB_OXI=<renderer.exe> adds Oxi)
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import feature_census as FC  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_mark_shadow_empty")
FONT = os.environ.get("PB_FONT", "ＭＳ ゴシック")
SH = ('<w14:shadow w14:blurRad="38100" w14:dist="22860" w14:dir="5400000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="t">'
      '<w14:srgbClr w14:val="000000"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:shadow>')


def slice_parts():
    src = [v["path"] for k, v in FC.load().items() if "5593c6a7" in k][0]
    zi = zipfile.ZipFile(src)
    parts = {n: zi.read(n) for n in zi.namelist()}
    x = parts["word/document.xml"].decode()
    i = x.find("宮城県柴田高等学校")
    a = x.rfind("<mc:AlternateContent>", 0, i)
    b = x.find("</mc:AlternateContent>", i) + len("</mc:AlternateContent>")
    return parts, x, re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", x[a:b], flags=re.S)


def para(runs, mark_extra=""):
    ppr = '<w:pPr><w:jc w:val="left"/>' + ("<w:rPr>%s</w:rPr>" % mark_extra if mark_extra else "") + "</w:pPr>"
    body = "".join('<w:r><w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr><w:t>%s</w:t></w:r>'
                   % (FONT, FONT, int(pt * 2), text) for pt, text in runs)
    return "<w:p>" + ppr + body + "</w:p>"


def build(name, mark_extra):
    parts, x, anchor = slice_parts()
    tc0 = anchor.find("<w:txbxContent>") + len("<w:txbxContent>")
    tc1 = anchor.find("</w:txbxContent>")
    content = para([(22, "部活の充実")]) + para([], mark_extra) + para([(22, "他にも有名")])
    anc = anchor[:tc0] + content + anchor[tc1:]
    nx = x[:x.find("<w:body>") + 8] + "<w:p><w:r>" + anc + "</w:r></w:p>" + x[x.rfind("<w:sectPr"):]
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
    ys = {}
    for b in pymupdf.open(pdf)[0].get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = "".join(c["c"] for c in s["chars"]).strip()
                if t:
                    ys.setdefault(t[:4], s["chars"][0]["origin"][1])
    return ys["他にも有名"[:4]] - ys["部活の充実"[:4]]


def pitch_oxi(exe, docx):
    import subprocess, json, tempfile
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    rows = {}
    for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"]:
        rows.setdefault(round(c["baseline"], 1), []).append(c)
    ys = {}
    for k in sorted(rows):
        tt = "".join(c["char"] for c in sorted(rows[k], key=lambda c: c["x"])).strip()
        if tt:
            ys.setdefault(tt[:4], k)
    return ys["他にも有名"[:4]] - ys["部活の充実"[:4]]


if __name__ == "__main__":
    import win32com.client
    arms = [("plain", ""), ("mark_shadow_10.5", SH), ("mark_22_noshadow", '<w:sz w:val="44"/>'),
            ("mark_shadow_22", '<w:sz w:val="44"/>' + SH), ("mark_shadow_26", '<w:sz w:val="52"/>' + SH)]
    base_w = base_o = None
    for name, extra in arms:
        p = build(name, extra); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        pw = pitch_word(pdf)
        po = pitch_oxi(os.environ["PB_OXI"], p) if os.environ.get("PB_OXI") else None
        if name == "plain":
            base_w, base_o = pw, po
        print("%-18s Word pitch %6.2f growth %+5.2f | Oxi %s" % (
            name, pw, pw - base_w, ("pitch %6.2f growth %+5.2f" % (po, po - base_o)) if po is not None else "-"), flush=True)
