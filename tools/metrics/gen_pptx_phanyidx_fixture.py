# -*- coding: utf-8 -*-
"""Build tests/fixtures/phanyidx_test.pptx.

A slide placeholder whose idx matches nothing: 4294967295, the sentinel
PowerPoint writes for an unset one. The layout has no placeholder at all and the
master's is `<p:ph type="title"/>` declaring sz="2400", while the master's
p:txStyles/p:bodyStyle says 1400. PowerPoint drew d24 slide 22's paragraph at
exactly 24.00pt, so the master PLACEHOLDER wins and the idx is not part of the
match.

The slide asks for `ctrTitle` and the master declares `title`, which are the
same slot, so this also pins the alias. The level carries a yellow
`a:highlight`, which d35's master title level uses for the white slab behind
BIG CONCEPT.

python-pptx cannot express this, so the parts are written by hand into a
minimal package.
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path("tests/fixtures/phanyidx_test.pptx").resolve()
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = f'xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"'


def sp(ph, body):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="2" name="ph"/><p:cNvSpPr/>'
            f'<p:nvPr>{ph}</p:nvPr></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="838200" y="838200"/>'
            f'<a:ext cx="4114800" cy="2743200"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/>{body}</p:txBody></p:sp>')


def tree(inner):
    return (f'<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
            f'<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{inner}'
            f'</p:spTree></p:cSld>')


LVL = ('<a:lstStyle><a:lvl1pPr><a:defRPr sz="2400" i="1">'
       '<a:solidFill><a:srgbClr val="112233"/></a:solidFill>'
       '<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>'
       '<a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></a:lstStyle>')

SLIDE = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         f'<p:sld {NS}>'
         + tree(sp('<p:ph idx="4294967295" type="ctrTitle"/>',
                   '<a:p><a:r><a:rPr lang="en"/><a:t>inherit me</a:t></a:r></a:p>'))
         + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')

LAYOUT = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<p:sldLayout {NS} type="blank">' + tree("")
          + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')

# The master's body placeholder is idx="1"; its lstStyle says 24pt while the
# txStyles bodyStyle says 14pt, so the two sources are distinguishable.
MASTER = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<p:sldMaster {NS}>'
          + tree('<p:sp><p:nvSpPr><p:cNvPr id="3" name="body"/><p:cNvSpPr/>'
                 '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
                 '<p:spPr/><p:txBody><a:bodyPr/>' + LVL + '<a:p/></p:txBody></p:sp>')
          + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"'
            ' accent2="accent2" accent3="accent3" accent4="accent4"'
            ' accent5="accent5" accent6="accent6" hlink="hlink"'
            ' folHlink="folHlink"/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/>'
            '</p:sldLayoutIdLst>'
            '<p:txStyles><p:titleStyle/><p:bodyStyle><a:lvl1pPr>'
            '<a:defRPr sz="1400"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/>'
            '</p:txStyles></p:sldMaster>')

PRES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:presentation {NS}>'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/>'
        '</p:sldMasterIdLst>'
        '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
        '<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>'
        '</p:presentation>')


def rels(items):
    body = "".join(
        f'<Relationship Id="{i}" Type="{t}" Target="{g}"/>' for i, t, g in items
    )
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PR}">{body}</Relationships>')


def main() -> None:
    O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   f'<Types xmlns="{CT}">'
                   '<Default Extension="rels" ContentType="application/'
                   'vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/ppt/presentation.xml" ContentType='
                   '"application/vnd.openxmlformats-officedocument.'
                   'presentationml.presentation.main+xml"/>'
                   '<Override PartName="/ppt/slides/slide1.xml" ContentType='
                   '"application/vnd.openxmlformats-officedocument.'
                   'presentationml.slide+xml"/>'
                   '<Override PartName="/ppt/slideLayouts/slideLayout1.xml"'
                   ' ContentType="application/vnd.openxmlformats-officedocument.'
                   'presentationml.slideLayout+xml"/>'
                   '<Override PartName="/ppt/slideMasters/slideMaster1.xml"'
                   ' ContentType="application/vnd.openxmlformats-officedocument.'
                   'presentationml.slideMaster+xml"/></Types>')
        z.writestr("_rels/.rels", rels([("rId1", O + "officeDocument",
                                         "ppt/presentation.xml")]))
        z.writestr("ppt/presentation.xml", PRES)
        z.writestr("ppt/_rels/presentation.xml.rels",
                   rels([("rId1", O + "slideMaster", "slideMasters/slideMaster1.xml"),
                         ("rId2", O + "slide", "slides/slide1.xml")]))
        z.writestr("ppt/slides/slide1.xml", SLIDE)
        z.writestr("ppt/slides/_rels/slide1.xml.rels",
                   rels([("rId1", O + "slideLayout",
                          "../slideLayouts/slideLayout1.xml")]))
        z.writestr("ppt/slideLayouts/slideLayout1.xml", LAYOUT)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                   rels([("rId1", O + "slideMaster",
                          "../slideMasters/slideMaster1.xml")]))
        z.writestr("ppt/slideMasters/slideMaster1.xml", MASTER)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",
                   rels([("rId1", O + "slideLayout",
                          "../slideLayouts/slideLayout1.xml")]))
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
