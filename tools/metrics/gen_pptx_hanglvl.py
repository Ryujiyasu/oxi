# -*- coding: utf-8 -*-
"""Probe: where does an unstated paragraph's hangingPunct come from?

252cbc8c slide 37/85: the paragraph says nothing, p:defaultTextStyle says
nothing, the master's p:bodyStyle levels say hangingPunct="1" -- PowerPoint
hangs the closing 。, Oxi pushes 「す。」 down. Each arm is its own file (the
master and defaultTextStyle are file-wide). A body placeholder (and a plain
text box) holds N full-width あ plus 。 in a box exactly N ems wide; COM
counts the lines: 1 = hung, 2 = pushed.

Arms (bodyStyle lvl1 / layout ph lstStyle lvl1 / defaultTextStyle lvl1):
  m1     1 / -  / -
  m0d1   0 / -  / 1
  m0l1   0 / 1  / -
  m1l0   1 / 0  / -
  none   - / -  / -
  d1     - / -  / 1
  d0     - / -  / 0
  m1d0   1 / -  / 0
"""
from __future__ import annotations

import os
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\hanglvl").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
N = 12
FS = 20
ARMS = {
    "m1": ("1", None, None),
    "m0d1": ("0", None, "1"),
    "m0l1": ("0", "1", None),
    "m1l0": ("1", "0", None),
    "none": (None, None, None),
    "d1": (None, None, "1"),
    "d0": (None, None, "0"),
    "m1d0": ("1", None, "0"),
}


def q(tag: str) -> str:
    return f"{{{A}}}{tag}"


def fill(tf) -> None:
    body = tf._txBody
    bp = body.find(q("bodyPr"))
    for k in ("lIns", "rIns", "tIns", "bIns"):
        bp.set(k, "0")
    bp.set("wrap", "square")
    for pel in body.findall(q("p")):
        body.remove(pel)
    p = etree.SubElement(body, q("p"))
    ppr = etree.SubElement(p, q("pPr"))
    ppr.set("marL", "0")
    ppr.set("indent", "0")
    etree.SubElement(ppr, q("buNone"))
    r = etree.SubElement(p, q("r"))
    rpr = etree.SubElement(r, q("rPr"))
    rpr.set("lang", "ja-JP")
    rpr.set("altLang", "en-US")
    rpr.set("sz", str(FS * 100))
    etree.SubElement(rpr, q("latin")).set("typeface", "ＭＳ ゴシック")
    etree.SubElement(rpr, q("ea")).set("typeface", "ＭＳ ゴシック")
    etree.SubElement(r, q("t")).text = "あ" * N + "。"


def set_lvl1(seg: str, v: str | None) -> str:
    if v is None:
        return re.sub(r' hangingPunct="[01]"', "", seg)
    seg = re.sub(r' hangingPunct="[01]"', "", seg)
    return re.sub(r"<a:lvl1pPr\b", f'<a:lvl1pPr hangingPunct="{v}"', seg, count=1)


def build(arm: str, m: str | None, l: str | None, d: str | None) -> Path:
    prs = Presentation()
    w = Emu(int(Pt(FS * N + 1)))
    s1 = prs.slides.add_slide(prs.slide_layouts[1])
    ph = s1.placeholders[1]
    ph.left, ph.top, ph.width, ph.height = Emu(914400), Emu(1828800), w, Emu(1828800)
    fill(ph.text_frame)
    box = s1.shapes.add_textbox(Emu(914400), Emu(4114800), w, Emu(914400))
    fill(box.text_frame)
    tmp = OUT / f"_{arm}.pptx"
    prs.save(tmp)
    dst = OUT / f"hanglvl_{arm}.pptx"
    with zipfile.ZipFile(tmp) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            name = it.filename
            if name in ("ppt/presentation.xml", "ppt/slideMasters/slideMaster1.xml",
                        "ppt/slideLayouts/slideLayout2.xml"):
                s = data.decode("utf-8")
                s = re.sub(r' hangingPunct="[01]"', "", s)
                if name == "ppt/presentation.xml":
                    i = s.find("<p:defaultTextStyle>"); j = s.find("</p:defaultTextStyle>")
                    s = s[:i] + set_lvl1(s[i:j], d) + s[j:]
                elif name.endswith("slideMaster1.xml"):
                    i = s.find("<p:bodyStyle>"); j = s.find("</p:bodyStyle>")
                    s = s[:i] + set_lvl1(s[i:j], m) + s[j:]
                    i = s.find("<p:otherStyle>"); j = s.find("</p:otherStyle>")
                    s = s[:i] + set_lvl1(s[i:j], None) + s[j:]
                elif l is not None:
                    # body placeholder (idx=1) of the layout: own lstStyle
                    k = s.find('idx="1"')
                    tb = s.find("<p:txBody>", k)
                    lst = s.find("<a:lstStyle/>", tb)
                    s = (s[:lst] + f'<a:lstStyle><a:lvl1pPr hangingPunct="{l}"/></a:lstStyle>'
                         + s[lst + len("<a:lstStyle/>"):])
                data = s.encode("utf-8")
            zo.writestr(it, data)
    tmp.unlink()
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    files = {arm: build(arm, *v) for arm, v in ARMS.items()}
    if "--no-com" in sys.argv:
        return
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        for arm, f in files.items():
            pres = app.Presentations.Open(str(f), WithWindow=False)
            sl = pres.Slides(1)
            rows = []
            for sh in sl.Shapes:
                if sh.HasTextFrame and sh.TextFrame.HasText:
                    tr = sh.TextFrame.TextRange
                    n = 0
                    while n < 20 and tr.Lines(n + 1, 1).Length > 0:
                        n += 1
                    rows.append((sh.Name, n, tr.ParagraphFormat.HangingPunctuation))
            pdf = str(f.with_suffix(".pdf"))
            pres.SaveAs(pdf, 32)
            pres.Close()
            print(arm, ARMS[arm], rows)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
