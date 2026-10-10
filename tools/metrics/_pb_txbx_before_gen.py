# -*- coding: utf-8 -*-
"""Does Word keep the first paragraph's space-before at the top of a text box?

ja/reference__5593c6a7: a text box whose first paragraph is pStyle Title
(spacing before=240) draws its 40pt line 12.2pt lower in Word than in Oxi.
One floating wps text box (the _pb_txbxfit_gen.py anchor, 120pt tall,
tIns 3.6pt) whose FIRST paragraph carries spacing before = B twips;
MS Gothic 12pt «壱», a second paragraph «弐».  Read: Word PDF baseline of
壱 / 弐 per arm; the shift from arm b0 is what before adds.

Arms: bB_gG  B in {0, 240, 480}, G = docGrid lines/360 (1) or none (0);
      cB_g1  as bB_g1 with bodyPr compatLnSpc="1";
      s240   before=240 from a paragraph STYLE (as in the real doc), grid 1.

    python tools/metrics/_pb_txbx_before_gen.py gen
    python tools/metrics/_pb_txbx_before_gen.py word
    python tools/metrics/_pb_txbx_before_gen.py oxi <renderer.exe>
"""
import json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402  (anchor / NS / package parts)

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "tests/fixtures/txbx_before"
ARMS = {"b0_g1": (0, 1, False), "b240_g1": (240, 1, False), "b480_g1": (480, 1, False),
        "b0_g0": (0, 0, False), "b240_g0": (240, 0, False), "s240_g1": (240, 1, True),
        # bodyPr compatLnSpc="1" (the real doc's box; Oxi's CLNSP law chain)
        "c0_g1": (0, 1, False), "c240_g1": (240, 1, False), "c480_g1": (480, 1, False)}

STYLES = T.STYLES.replace(
    "</w:styles>",
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:before="240" w:after="0"/></w:pPr></w:style></w:styles>')


def paras(before, from_style):
    ppr1 = ('<w:pPr><w:pStyle w:val="Title"/></w:pPr>' if from_style
            else f'<w:pPr><w:spacing w:before="{before}" w:after="0"/></w:pPr>')
    return (f'<w:p>{ppr1}<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>壱</w:t></w:r></w:p>'
            '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>弐</w:t></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (before, grid, from_style) in ARMS.items():
        a = T.anchor(120, 1, False)
        if arm.startswith("c"):
            a = a.replace(' anchor="t"', ' anchor="t" compatLnSpc="1"')
        a = re.sub(r"<w:txbxContent>.*</w:txbxContent>", "<w:txbxContent>" + paras(before, from_style) + "</w:txbxContent>", a, flags=re.S)
        body = "<w:p>" + a + "</w:p>" + "".join(f"<w:p><w:r><w:t>L{i:02d}</w:t></w:r></w:p>" for i in range(1, 9))
        grid_xml = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ""
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
               ' w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="425"/>' + grid_xml + "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", T.CT); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", T.DRELS); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/document.xml", doc)
    print("ok", list(ARMS))


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        ys = {}
        for b in fitz.open(OUT / f"{arm}.pdf")[0].get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"] in "壱弐" and c["c"] not in ys:
                            ys[c["c"]] = round(c["origin"][1], 2)
        res[arm] = ys
        print(arm, ys, flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")


def oxi(exe):
    sys.path.insert(0, str(REPO / "tools/metrics"))
    import line_census as L
    res = {}
    for arm in ARMS:
        ys = {}
        for o in L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"):
            for ch in "壱弐":
                if o["t"].startswith(ch) and ch not in ys:
                    ys[ch] = round(o["y"], 2)
        res[arm] = ys
        print(arm, ys, flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
