# -*- coding: utf-8 -*-
"""Build tests/fixtures/grpfliprot_test.pptx.

A layout whose group is rot=-90 flipH=1 and holds another group that is also
rot=-90 flipH=1. Component-wise accumulation makes that -180 with the flips
cancelling; the real composition is R(-90) F R(-90) F = R(-90) R(+90) = the
IDENTITY, because F R(t) = R(-t) F. d19's 29 layout pencils are exactly this
shape and Oxi drew every one of them upside down.

The slide is empty, so the only shape in the result is the layout's.
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path("tests/fixtures/grpfliprot_test.pptx").resolve()
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = f'xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"'


def grp(rot, off, ext, choff, chext, inner):
    return (f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="g"/><p:cNvGrpSpPr/>'
            f'<p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="{rot}" flipH="1">'
            f'<a:off x="{off[0]}" y="{off[1]}"/><a:ext cx="{ext[0]}" cy="{ext[1]}"/>'
            f'<a:chOff x="{choff[0]}" y="{choff[1]}"/>'
            f'<a:chExt cx="{chext[0]}" cy="{chext[1]}"/></a:xfrm></p:grpSpPr>'
            f'{inner}</p:grpSp>')


LEAF = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="leaf"/><p:cNvSpPr/><p:nvPr/>'
        '</p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/>'
        '<a:ext cx="914400" cy="1828800"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>')

# Both groups map their child box onto itself, so only the rotations compose.
BOX = (914400, 914400, 914400, 1828800)
INNER = grp(-5400000, (BOX[0], BOX[1]), (BOX[2], BOX[3]),
            (BOX[0], BOX[1]), (BOX[2], BOX[3]), LEAF)
OUTER = grp(-5400000, (BOX[0], BOX[1]), (BOX[2], BOX[3]),
            (BOX[0], BOX[1]), (BOX[2], BOX[3]), INNER)


def tree(inner):
    return (f'<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
            f'<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{inner}'
            f'</p:spTree></p:cSld>')


SLIDE = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld {NS}>'
         + tree("") + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')
LAYOUT = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<p:sldLayout {NS} type="blank">' + tree(OUTER)
          + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')
MASTER = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<p:sldMaster {NS}>' + tree("")
          + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"'
            ' accent2="accent2" accent3="accent3" accent4="accent4"'
            ' accent5="accent5" accent6="accent6" hlink="hlink"'
            ' folHlink="folHlink"/><p:sldLayoutIdLst>'
            '<p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
            '<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/>'
            '</p:txStyles></p:sldMaster>')
PRES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:presentation {NS}><p:sldMasterIdLst>'
        '<p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
        '<p:sldSz cx="9144000" cy="6858000"/>'
        '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>')


def rels(items):
    body = "".join(f'<Relationship Id="{i}" Type="{t}" Target="{g}"/>'
                   for i, t, g in items)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PR}">{body}</Relationships>')


def main() -> None:
    O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   f'<Types xmlns="{CT}"><Default Extension="rels" ContentType='
                   '"application/vnd.openxmlformats-package.relationships+xml"/>'
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
                   rels([("rId1", O + "slideMaster",
                          "slideMasters/slideMaster1.xml"),
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
