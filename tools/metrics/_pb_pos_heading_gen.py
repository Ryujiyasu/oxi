# -*- coding: utf-8 -*-
"""Why is the HG創英角ﾎﾟｯﾌﾟ体 52pt heading of blind-G ja/policies__20069856 set 1.53pt
too high in Oxi when its neighbours in the same text box are within 0.35?

The `_pb_txbx_cjk_base_gen.py` arms (single wps box, HGP face, 52pt, anchor ctr)
already match, so the cause is something only this document has. Faithful
slices of the real package, one property changed per arm:

    asis        the document as it is
    nopos       docDefaults <w:position w:val="10"/> removed (every run is raised 5pt)
    pos20       docDefaults position 20 (10pt raise)
    hgp         the two 52pt headings' face HG創英角ﾎﾟｯﾌﾟ体 -> HGP創英角ﾎﾟｯﾌﾟ体
    sz72        the two headings at 36pt instead of 52pt
    lnwnone / lnw6350 / lnw76200   the two title members' `a:ln w` (asis 38100 = 3pt, noFill)
                removed / 0.5pt / 6pt -- the outline inset (archive (57), w/2 with noFill too)

Read: baselines of «意思表示のポイント» / «意思表示のいわか» and of the line before
and after them (Word PDF; Oxi glyph dump via line_census).

    python tools/metrics/_pb_pos_heading_gen.py          (PB_OXI=<renderer.exe> adds Oxi)
"""
import os, re, sys, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
import line_census as LC  # noqa: E402

SRC = REPO / "pipeline_data/docx_corpus/ja/policies/20069856e07302d0.docx"
OUT = REPO / "pipeline_data" / "_pb_pos_heading"
HEAD = "HG創英角ﾎﾟｯﾌﾟ体"


def variant(name):
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    doc = parts["word/document.xml"].decode("utf-8")
    sty = parts["word/styles.xml"].decode("utf-8")
    if name == "nopos":
        sty = sty.replace('<w:position w:val="10"/>', "", 1)
    elif name == "pos20":
        sty = sty.replace('<w:position w:val="10"/>', '<w:position w:val="20"/>', 1)
    elif name == "hgp":
        doc = doc.replace(f'"{HEAD}"', '"HGP創英角ﾎﾟｯﾌﾟ体"')
    elif name.startswith("lnw"):
        # the two title members' outline width (asis 38100 EMU = 3pt, noFill)
        w = name[3:]
        doc = doc.replace('<a:ln w="38100"><a:noFill/></a:ln>',
                          '<a:ln><a:noFill/></a:ln>' if w == "none" else f'<a:ln w="{w}"><a:noFill/></a:ln>')
    elif name == "sz72":
        doc = doc.replace('<w:sz w:val="104"/>', '<w:sz w:val="72"/>')
    parts["word/document.xml"] = doc.encode("utf-8")
    parts["word/styles.xml"] = sty.encode("utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


def pick(lines):
    keys = ("いしひょうじ", "意思表示のポイント", "覚えておこう", "意思表示のいわか", "①い賛成")
    got = {}
    for k in keys:
        ys = sorted(l["y"] for l in lines if l["t"].startswith(k))
        got[k[:7]] = ys
    return got


if __name__ == "__main__":
    import win32com.client
    exe = os.environ.get("PB_OXI")
    for name in (os.environ.get("PB_ARMS", "asis,nopos,pos20,hgp,sz72,lnwnone,lnw6350,lnw76200").split(",")):
        p = variant(name); pdf = p.with_suffix(".pdf")
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(str(p), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
        finally:
            w.Quit()
        W = pick(LC.word_lines(pdf))
        O = pick(LC.oxi_lines(exe, p) or []) if exe else {}
        print(f"{name:6}", "  ".join(f"{k} W {' '.join('%.2f' % y for y in W[k])}"
                                     + (f" O {' '.join('%.2f' % y for y in O.get(k, []))}" if exe else "") for k in W), flush=True)
