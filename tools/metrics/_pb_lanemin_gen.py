# -*- coding: utf-8 -*-
"""Minimum lane beside a square-wrap text box: how narrow a right-hand lane does
Word still flow a line into?  (09422f63's heading «くしゃみは…» sits BELOW a
box that leaves a 32pt lane, where Oxi's 30pt threshold flows it in.)

Host package: ja/educational 09422f63 (fonts, grid). Body per arm: a paragraph
carrying the host's own wrapSquare text box (anchor #2, cx set so that the lane
right of it is L pt), then one 16pt line «くしゃみは時速何キロで出るかしっていますか？»,
with or without a leading tab (the heading has one; tab stop 1965tw).
Read from the Word PDF: the first glyph «く» origin -> beside (y inside the box's
band) or below.

    python tools/metrics/_pb_lanemin_gen.py            # Word only
    python tools/metrics/_pb_lanemin_gen.py <oxi.exe>  # + Oxi
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/educational/09422f63e991d48f.docx"
OUT = REPO / "tests/fixtures/lanemin"
LANES = [int(v) for v in os.environ.get("LM_LANES", "20,26,30,32,36,42,50,60").split(",")]
TABS = [False, True]
CONTENT_W = 595.3 - 2 * 42.55  # A4, margins 851tw
FONT = '<w:rFonts w:ascii="HG丸ｺﾞｼｯｸM-PRO" w:eastAsia="HG丸ｺﾞｼｯｸM-PRO" w:hAnsi="HG丸ｺﾞｼｯｸM-PRO" w:hint="eastAsia"/>'
TEXT = "くしゃみは時速何キロで出るかしっていますか？"


def name(L, tab):
    return f"L{L}_{'tab' if tab else 'notab'}"


def host_anchor(doc):
    anchors = re.findall(r"<wp:anchor\b.*?</wp:anchor>", doc, re.S)
    a = next(a for a in anchors if "wrapSquare" in a and "txbxContent" in a and "wps:wsp" in a)
    return a


def body(doc, L, tab):
    a = host_anchor(doc)
    cx = int((CONTENT_W - L - 9.0) * 12700)  # lane L after the 9pt distR
    a = re.sub(r'<wp:extent cx="\d+"', f'<wp:extent cx="{cx}"', a, count=1)
    a = re.sub(r'<a:ext cx="\d+"', f'<a:ext cx="{cx}"', a, count=1)
    a = re.sub(r'<wp:positionH relativeFrom="\w+">.*?</wp:positionH>',
               '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>', a, count=1, flags=re.S)
    a = re.sub(r'<wp:positionV relativeFrom="\w+">.*?</wp:positionV>',
               '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>', a, count=1, flags=re.S)
    a = re.sub(r'<wp:wrapSquare [^>]*/>', '<wp:wrapSquare wrapText="bothSides"/>', a, count=1)
    host_p = ('<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr><w:r><w:drawing>' + a + '</w:drawing></w:r></w:p>')
    rpr = f'<w:rPr>{FONT}<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
    tabx = '<w:r>' + rpr + '<w:tab/></w:r>' if tab else ''
    test_p = ('<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1965"/></w:tabs><w:snapToGrid w:val="0"/>'
              f'<w:rPr>{FONT}<w:sz w:val="32"/></w:rPr></w:pPr>{tabx}<w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>')
    filler = ''.join(f'<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr><w:r>{rpr}<w:t>本文{i}行目のテキストです。</w:t></w:r></w:p>' for i in range(4))
    return host_p + test_p + filler


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    b0 = doc.index("<w:body>") + len("<w:body>")
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    for L in LANES:
        for tab in TABS:
            xml = doc[:b0] + body(doc, L, tab) + sect + "</w:body></w:document>"
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                    zout.writestr(copy.copy(item), data)
            (OUT / f"{name(L, tab)}.docx").write_bytes(buf.getvalue())


def word():
    import fitz, win32com.client
    res = {}
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for L in LANES:
            for tab in TABS:
                n = name(L, tab)
                tmp = os.path.join(tempfile.mkdtemp(), n + ".docx"); shutil.copy(OUT / f"{n}.docx", tmp)
                d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
                pdf = tmp[:-5] + ".pdf"; d.SaveAs2(pdf, 17); d.Close(0)
                page = fitz.open(pdf)[0]
                got = None
                for b in page.get_text("rawdict")["blocks"]:
                    for l in b.get("lines", []):
                        for s in l["spans"]:
                            for c in s["chars"]:
                                if c["c"] == "く" and s["size"] > 15 and got is None:
                                    got = (round(c["origin"][0], 2), round(c["origin"][1], 2))
                res[n] = got
    finally:
        w.Quit()
    return res


def oxi(exe):
    res = {}
    for L in LANES:
        for tab in TABS:
            n = name(L, tab); t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
            subprocess.run([exe, str(OUT / f"{n}.docx"), os.path.join(t, "p"), "110", "--dump-glyphs=" + out], capture_output=True)
            g = json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]
            c = next((x for x in g if x["char"] == "く" and x["font_size"] > 15), None)
            res[n] = (round(c["x"], 2), round(c["top"] + 0.859 * c["font_size"], 2)) if c else None
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    gen()
    W = word()
    O = oxi(os.path.abspath(sys.argv[1])) if len(sys.argv) > 1 else {}
    for L in LANES:
        for tab in TABS:
            n = name(L, tab)
            print(f"{n:12} Word く at {W[n]}   Oxi {O.get(n)}")
