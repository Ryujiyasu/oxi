# -*- coding: utf-8 -*-
"""Probe: does a plain text box take its size from p:defaultTextStyle or from
the master's p:txStyles/p:otherStyle?

abde73e7 says defaultTextStyle sz=1800 and otherStyle sz=1350; PowerPoint draws
its unsized text boxes at 18pt and Oxi at 13.5pt. The arms set the two to
different sizes (and colours) so the export says which one wins, per level.
"""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\otherstyle").resolve()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def q(tag: str) -> str:
    return f"{{{A}}}{tag}"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    box = slide.shapes.add_textbox(Emu(914400), Emu(914400), Emu(6400800), Emu(3000000))
    body = box.text_frame._txBody
    for pel in body.findall(q("p")):
        body.remove(pel)
    for lvl, text in ((0, "Level one HHHH"), (1, "Level two HHHH")):
        p = etree.SubElement(body, q("p"))
        if lvl:
            etree.SubElement(p, q("pPr")).set("lvl", str(lvl))
        r = etree.SubElement(p, q("r"))
        etree.SubElement(r, q("rPr")).set("lang", "en-US")
        etree.SubElement(r, q("t")).text = text
    tmp = OUT / "_tmp.pptx"
    prs.save(tmp)
    dst = OUT / "otherstyle.pptx"
    with zipfile.ZipFile(tmp) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename == "ppt/presentation.xml":
                s = data.decode("utf-8")
                # defaultTextStyle: every level 24pt, red
                i = s.find("<p:defaultTextStyle>"); j = s.find("</p:defaultTextStyle>")
                seg = re.sub(r'sz="\d+"', 'sz="2400"', s[i:j])
                seg = re.sub(r'<a:schemeClr val="tx1"/>', '<a:srgbClr val="FF0000"/>', seg)
                s = s[:i] + seg + s[j:]
                data = s.encode("utf-8")
            elif it.filename == "ppt/slideMasters/slideMaster1.xml":
                s = data.decode("utf-8")
                i = s.find("<p:otherStyle>"); j = s.find("</p:otherStyle>")
                seg = re.sub(r'sz="\d+"', 'sz="1200"', s[i:j])
                seg = re.sub(r'<a:schemeClr val="tx1"/>', '<a:srgbClr val="0000FF"/>', seg)
                s = s[:i] + seg + s[j:]
                data = s.encode("utf-8")
            zo.writestr(it, data)
    tmp.unlink()
    print("wrote", dst)


if __name__ == "__main__":
    main()
