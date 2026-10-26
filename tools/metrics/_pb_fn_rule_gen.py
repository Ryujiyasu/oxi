# -*- coding: utf-8 -*-
"""Where inside the separator paragraph does Word draw the footnote rule?

blind-G ja/educational__20aa108d: Word's rule is at 649.80, Oxi's at 642.16
(Oxi draws it at the top of the separator allocation). Faithful slices of the
real package; only the separator footnote's paragraph changes per arm:
    asis       <w:p><w:r><w:separator/></w:r></w:p> (Normal, 10.5pt, lines 360)
    nogrid     the section's docGrid removed
    sz20       the separator run at 20pt
    sz8        the separator run at 8pt
    exact30    the separator paragraph line=600 exact
    before12   the separator paragraph before=240
    after12    the separator paragraph after=240
    nogrid_sz20 / nogrid_exact30 / nogrid_before12 / nogrid_after12
    arial      the separator run in Arial (rFonts ascii/hAnsi)
    msmin      the separator run in ＭＳ 明朝 (ascii/hAnsi/eastAsia)
    eagothic   the separator run with only eastAsia=ＭＳ ゴシック
    times20    the separator run in Times New Roman 20pt
    yumin / meiryo / calibri  the separator run in 游明朝 / メイリオ / Calibri (all four rFonts)
    double     the separator paragraph line=480 auto
    atleast30  the separator paragraph line=600 atLeast
Read: the rule (the 144pt-wide drawing) top and bottom, the note's first
baseline («本»).

    python tools/metrics/_pb_fn_rule_gen.py     (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import json, os, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/educational/20aa108db65e9ee3.docx"
OUT = REPO / "pipeline_data" / "_pb_fn_rule"
SEP = '<w:r><w:separator/></w:r>'
ARMS = ["asis", "sz20", "sz8", "exact30", "before12", "after12",
        "nogrid", "nogrid_sz20", "nogrid_exact30", "nogrid_before12", "nogrid_after12",
        "arial", "msmin", "eagothic", "times20", "double", "atleast30", "yumin", "meiryo", "calibri"]
FONTS = {"yumin": '<w:rFonts w:ascii="游明朝" w:eastAsia="游明朝" w:hAnsi="游明朝"/>',
         "meiryo": '<w:rFonts w:ascii="メイリオ" w:eastAsia="メイリオ" w:hAnsi="メイリオ"/>',
         "calibri": '<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/>',
         "arial": '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>',
         "msmin": '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>',
         "eagothic": '<w:rFonts w:eastAsia="ＭＳ ゴシック"/>',
         "times20": '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="40"/><w:szCs w:val="40"/>'}


def variant(name):
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    doc = parts["word/document.xml"].decode("utf-8")
    fn = parts["word/footnotes.xml"].decode("utf-8")
    i = fn.find('w:type="separator"')
    j = fn.find("</w:footnote>", i)
    sep = fn[i:j]
    assert SEP in sep
    arm = name.replace("nogrid_", "") if name != "nogrid" else "asis"
    ppr = {"exact30": '<w:pPr><w:spacing w:line="600" w:lineRule="exact"/></w:pPr>',
           "before12": '<w:pPr><w:spacing w:before="240"/></w:pPr>',
           "after12": '<w:pPr><w:spacing w:after="240"/></w:pPr>',
           "double": '<w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>',
           "atleast30": '<w:pPr><w:spacing w:line="600" w:lineRule="atLeast"/></w:pPr>'}.get(arm, "")
    run = SEP
    if arm in ("sz20", "sz8"):
        v = int(arm[2:]) * 2
        run = f'<w:r><w:rPr><w:sz w:val="{v}"/><w:szCs w:val="{v}"/></w:rPr><w:separator/></w:r>'
        ppr = f'<w:pPr><w:rPr><w:sz w:val="{v}"/><w:szCs w:val="{v}"/></w:rPr></w:pPr>'
    if arm in FONTS:
        run = f'<w:r><w:rPr>{FONTS[arm]}</w:rPr><w:separator/></w:r>'
        ppr = f'<w:pPr><w:rPr>{FONTS[arm]}</w:rPr></w:pPr>'
    k = sep.find("<w:r>")
    new = sep[:k] + ppr + run + sep[k + len(SEP):]
    fn = fn[:i] + new + fn[j:]
    if name.startswith("nogrid"):
        doc = doc.replace('<w:docGrid w:type="lines" w:linePitch="360"/>', "")
    parts.update({"word/document.xml": doc.encode("utf-8"), "word/footnotes.xml": fn.encode("utf-8")})
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


def read_word(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    base = None
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"] == "本" and c["origin"][1] > 640 and base is None:
                        base = c["origin"][1]
    rule = [(d["rect"].y0, d["rect"].y1) for d in pg.get_drawings()
            if abs(d["rect"].width - 144) < 1 and d["rect"].y0 > 450]
    return base, rule[0] if rule else None


def read_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json"); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    subprocess.run([exe, str(docx), os.path.join(t, "q"), "--dump-layout=" + lay], capture_output=True)
    base = None
    for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"]:
        if c["char"] == "本" and c["baseline"] > 640 and base is None:
            base = c["baseline"]
    rule = [(e["y"], e["y"] + e.get("h", 0)) for e in json.load(open(lay, encoding="utf-8"))["pages"][0]["elements"]
            if abs(e.get("w", 0) - 144) < 1 and e["y"] > 450]
    return base, rule[0] if rule else None


def fmt(v):
    base, rule = v
    if base is None or rule is None:
        return f"note {base} rule {rule}"
    return f"note {base:7.2f} rule {rule[0]:7.2f}-{rule[1]:7.2f} (note-rule {base - rule[0]:6.2f})"


if __name__ == "__main__":
    exes = [e for e in os.environ.get("PB_OXI", "").split(";") if e]
    for name in os.environ.get("PB_ARMS", ",".join(ARMS)).split(","):
        p = variant(name); pdf = p.with_suffix(".pdf")
        if not pdf.exists() or os.environ.get("PB_REWORD"):
            import win32com.client
            w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
            try:
                d = w.Documents.Open(str(p), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
            finally:
                w.Quit()
        print(f"{name:16} W {fmt(read_word(pdf))}" + "".join(f"   O {fmt(read_oxi(e, p))}" for e in exes), flush=True)
