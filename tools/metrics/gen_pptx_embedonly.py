# -*- coding: utf-8 -*-
"""Probe: which face does PowerPoint draw a b=0 run with when the deck embeds
only SOME slots of a typeface the machine's cloud cache also serves?

d60 embeds Poppins bold + boldItalic only and PowerPoint draws its b=0 Poppins
text with the embedded Bold; d46 embeds Open Sans bold + boldItalic only and
PowerPoint's b=0 Open Sans comes from the cloud Regular. Each arm here is one
deck with one typeface, its parts copied byte for byte from those decks, and
two runs (b=0, b=1). Output: the PDF font each run was drawn with (embedded
parts come out numbered, machine fonts by name) and that font's weight.
"""
from __future__ import annotations

import io
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "pipeline_data" / "pptx_probes" / "embedonly"
DEV = ROOT / "pipeline_data" / "pptx_benchmark" / "dev" / "pptx"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
FONT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"


def parts(deck_prefix: str, typeface: str) -> dict[str, bytes]:
    f = next(iter(sorted(DEV.glob(deck_prefix + "*.pptx")) + sorted((DEV.parent.parent / "pptx").glob(deck_prefix + "*.pptx"))))
    z = zipfile.ZipFile(f)
    p = z.read("ppt/presentation.xml").decode("utf-8")
    rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    out = {}
    for m in re.finditer(r"<p:embeddedFont><p:font ([^>]*)/>(.*?)</p:embeddedFont>", p):
        if f'typeface="{typeface}"' not in m.group(1):
            continue
        for slot, rid in re.findall(r'<p:(\w+) r:id="(\w+)"/>', m.group(2)):
            t = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(
                r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
            out[slot] = z.read("ppt/" + t.group(1))
    return out


POP = parts("d60_", "Poppins")
POP46 = parts("d46_", "Poppins")
OS = parts("d46_", "Open Sans")
ARMS = {
    "poppins_bold_only": ("Poppins", {"bold": POP["bold"]}),
    "poppins_bold_bi": ("Poppins", {"bold": POP["bold"], "boldItalic": POP["boldItalic"]}),
    "poppins_reg_bold": ("Poppins", {"regular": POP46["regular"], "bold": POP["bold"]}),
    "opensans_bold_only": ("Open Sans", {"bold": OS["bold"]}),
    "opensans_bold_bi": ("Open Sans", {"bold": OS["bold"], "boldItalic": OS["boldItalic"]}),
    "none_poppins": ("Poppins", {}),
    "opensans110_bold_only": ("Open Sans", {"bold": parts("d35_", "Open Sans")["bold"]}),
    "opensans3003_bold_only": ("Open Sans", {"bold": parts("d53_", "Open Sans")["bold"]}),
    "montserrat72_bold_only": ("Montserrat", {"bold": parts("d11_", "Montserrat")["bold"]}),
    "montserrat8_bold_only": ("Montserrat", {"bold": parts("07__", "Montserrat")["bold"]}),
    "lato_bold_only": ("Lato", {"bold": parts("d31_", "Lato")["bold"]}),
    "nunito_bold_only": ("Nunito", {"bold": parts("d02_", "Nunito")["bold"]}),
    "roboto_bold_only": ("Roboto", {"bold": parts("d58_", "Roboto")["bold"]}),
}
ONLY = sys.argv[1:]


def q(ns: str, t: str) -> str:
    return f"{{{ns}}}{t}"


def emu(pt: float) -> Emu:
    return Emu(int(round(pt * 12700)))


def build(name: str, face: str, fparts: dict[str, bytes]) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(960), emu(540)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for k, bold in enumerate((False, True)):
        box = slide.shapes.add_textbox(emu(20), emu(20 + k * 120), emu(900), emu(100))
        body = box.text_frame._txBody
        for pel in body.findall(q(A, "p")):
            body.remove(pel)
        p = etree.SubElement(body, q(A, "p"))
        r = etree.SubElement(p, q(A, "r"))
        rpr = etree.SubElement(r, q(A, "rPr"))
        rpr.set("lang", "en-US")
        rpr.set("sz", "4000")
        rpr.set("b", "1" if bold else "0")
        etree.SubElement(rpr, q(A, "latin")).set("typeface", face)
        etree.SubElement(r, q(A, "t")).text = ("BOLD " if bold else "PLAIN ") + "Hamburgefonstiv"
    tmp = OUT / f"{name}_raw.pptx"
    prs.save(tmp)
    if not fparts:
        dst = OUT / f"{name}.pptx"
        tmp.replace(dst)
        return dst
    # add the font parts + presentation.xml embeddedFontLst
    zin = zipfile.ZipFile(tmp)
    dst = OUT / f"{name}.pptx"
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    rel_xml = zin.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    pres_xml = zin.read("ppt/presentation.xml").decode("utf-8")
    ct = zin.read("[Content_Types].xml").decode("utf-8")
    slots = []
    for i, (slot, data) in enumerate(fparts.items()):
        rid = f"rIdF{i}"
        rel_xml = rel_xml.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" Type="{FONT_REL}" Target="fonts/font{i}.fntdata"/></Relationships>')
        slots.append((slot, rid))
        zout.writestr(f"ppt/fonts/font{i}.fntdata", data)
    order = ["regular", "bold", "italic", "boldItalic"]
    slots.sort(key=lambda s: order.index(s[0]))
    lst = (f'<p:embeddedFontLst><p:embeddedFont><p:font typeface="{face}"/>'
           + "".join(f'<p:{s} r:id="{rid}"/>' for s, rid in slots)
           + "</p:embeddedFont></p:embeddedFontLst>")
    # embeddedFontLst goes after sldSz/notesSz
    pres_xml = re.sub(r"(<p:notesSz[^>]*/>)", r"\1" + lst, pres_xml, count=1)
    # PowerPoint reads embeddedFontLst only when the deck says it embeds.
    pres_xml = pres_xml.replace("<p:presentation ", '<p:presentation embedTrueTypeFonts="1" ', 1)
    if 'Extension="fntdata"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)
    for item in zin.infolist():
        if item.filename == "ppt/_rels/presentation.xml.rels":
            zout.writestr(item, rel_xml)
        elif item.filename == "ppt/presentation.xml":
            zout.writestr(item, pres_xml)
        elif item.filename == "[Content_Types].xml":
            zout.writestr(item, ct)
        else:
            zout.writestr(item, zin.read(item.filename))
    zout.close()
    zin.close()
    tmp.unlink()
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    files = {n: build(n, face, fp) for n, (face, fp) in ARMS.items() if not ONLY or n in ONLY}
    import pymupdf
    import win32com.client
    from fontTools.ttLib import TTFont
    import subprocess
    import time
    # One PowerPoint PROCESS per arm: an embedded family loaded by one deck
    # stays mapped after it closes (a later embed-free Poppins deck drew in
    # Calibri), so arms sharing a process measure each other.
    for n, f in files.items():
        subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
        time.sleep(2)
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            pres = app.Presentations.Open(str(f), WithWindow=False)
            pres.SaveAs(str(f.with_suffix(".pdf")), 32)
            pres.Close()
        finally:
            app.Quit()
    for n, f in files.items():
        doc = pymupdf.open(f.with_suffix(".pdf"))
        page = doc[0]
        xref_of = {x[4]: x for x in page.get_fonts()}
        res = []
        for b in page.get_text("dict")["blocks"]:
            for line in b.get("lines", []):
                for sp in line["spans"]:
                    if not sp["text"].strip():
                        continue
                    font = sp["font"]
                    w = None
                    for x in page.get_fonts():
                        if x[3].endswith(font) or x[3].split("+")[-1] == font:
                            try:
                                t = TTFont(io.BytesIO(doc.extract_font(x[0])[3]))
                                w = t["OS/2"].usWeightClass
                            except Exception:
                                pass
                    res.append((sp["text"].split()[0], font, w, round(sp["bbox"][2] - sp["bbox"][0], 1)))
        print(n.ljust(20), res)


if __name__ == "__main__":
    main()
