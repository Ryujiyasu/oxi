# -*- coding: utf-8 -*-
"""Probe: which face does PowerPoint draw a run's East Asian characters in?

640d6e4e slide 88 runs state only `<a:latin typeface="ＭＳ"/>` (a face that does
not exist) over Japanese text; PowerPoint draws the kana in the inherited
Meiryo UI, Oxi in the substitute for "ＭＳ". Each arm is a text box holding
"Abc あいう漢字"; the PDF's span fonts say which face drew which script.

File A (python-pptx theme: minor latin Calibri, minor ea empty):
  a1  run latin Arial
  a2  run latin Arial + ea ＭＳ 明朝
  a3  run latin ＭＳ 明朝 only
  a4  run with no face at all
File B: the same, with the theme's minor ea set to ＭＳ ゴシック.
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\eaface").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
ARMS = [
    ("a1", "Arial", None),
    ("a2", "Arial", "ＭＳ 明朝"),
    ("a3", "ＭＳ 明朝", None),
    ("a4", None, None),
]


def q(t: str) -> str:
    return f"{{{A}}}{t}"


def build(tag: str, theme_ea: str | None) -> Path:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for k, (label, latin, ea) in enumerate(ARMS):
        box = slide.shapes.add_textbox(Emu(457200), Emu(457200 + k * 1100000), Emu(6400000), Emu(900000))
        body = box.text_frame._txBody
        for pel in body.findall(q("p")):
            body.remove(pel)
        p = etree.SubElement(body, q("p"))
        r = etree.SubElement(p, q("r"))
        rpr = etree.SubElement(r, q("rPr"))
        rpr.set("lang", "ja-JP")
        rpr.set("sz", "2800")
        if latin:
            etree.SubElement(rpr, q("latin")).set("typeface", latin)
        if ea:
            etree.SubElement(rpr, q("ea")).set("typeface", ea)
        etree.SubElement(r, q("t")).text = f"{label} Abc あいう漢字"
    tmp = OUT / f"_{tag}.pptx"
    prs.save(tmp)
    dst = OUT / f"eaface_{tag}.pptx"
    with zipfile.ZipFile(tmp) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if theme_ea and re.match(r"ppt/theme/theme1\.xml", it.filename):
                s = data.decode("utf-8")
                i = s.find("<a:minorFont>")
                j = s.find("</a:minorFont>", i)
                seg = re.sub(r'<a:ea typeface="[^"]*"', f'<a:ea typeface="{theme_ea}"', s[i:j], count=1)
                data = (s[:i] + seg + s[j:]).encode("utf-8")
            zo.writestr(it, data)
    tmp.unlink()
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    files = [build("a", None), build("b", "ＭＳ ゴシック")]
    import pymupdf
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        for f in files:
            pres = app.Presentations.Open(str(f), WithWindow=False)
            pdf = f.with_suffix(".pdf")
            pres.SaveAs(str(pdf), 32)
            pres.Close()
            for b in pymupdf.open(pdf)[0].get_text("dict")["blocks"]:
                for line in b.get("lines", []):
                    print(f.stem, " | ".join(f"{s['font']}:{s['text']}" for s in line["spans"]))
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
