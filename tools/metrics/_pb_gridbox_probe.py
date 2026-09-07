# -*- coding: utf-8 -*-
"""Where does Word put a line whose font is TALLER than one grid line?

On a typed docGrid a line takes ceil(need / pitch) grid lines; this asks where
inside that block the glyphs land.  One paragraph per size, each a short CJK
run, into a grid host document; the PDF gives every line's baseline, and the
grid lines are the page's top margin plus k * linePitch.

Usage: python tools/metrics/_pb_gridbox_probe.py [host-docx-stem]
"""
import glob
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
stem = sys.argv[1] if len(sys.argv) > 1 else "1ec1"
src = glob.glob("tools/golden-test/documents/docx/%s*.docx" % stem)[0]
z = zipfile.ZipFile(src)
doc = z.read("word/document.xml").decode("utf-8")
head = doc[: doc.index("<w:body>") + len("<w:body>")]
sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)[-1]
sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
SIZES = [float(x) for x in os.environ.get("GB_SIZES", "9,10.5,12,14,16,18,20,24,28").split(",")]
FONT = os.environ.get("GB_FONT", "ＭＳ ゴシック")
paras = []
for i, size in enumerate(SIZES):
    rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (FONT, FONT, FONT, round(size * 2), round(size * 2)))
    # GB_RULE=exact/atLeast + GB_LINE=<pt> puts the arms on an explicit line rule,
    # which is the OTHER placement regime (1ec1's body is line=360 exact).
    rule = os.environ.get("GB_RULE")
    line_pt = float(os.environ.get("GB_LINE", "18"))
    ppr = ('<w:spacing w:line="%d" w:lineRule="%s"/>' % (round(line_pt * 20), rule)) if rule else ""
    paras.append('<w:p><w:pPr>%s</w:pPr><w:r>%s<w:t>%s%02d日本語の行</w:t></w:r></w:p>'
                 % (ppr, rpr, chr(ord("A") + i), i))
out = os.path.join("pipeline_data", "_pb_gridbox_probe.docx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        if item.filename == "word/document.xml":
            o.writestr(item, (head + "".join(paras) + sect + "</w:body></w:document>").encode("utf-8"))
        elif item.filename.startswith("word/header") or item.filename.startswith("word/footer"):
            continue
        else:
            data = z.read(item.filename)
            if item.filename == "word/_rels/document.xml.rels":
                data = re.sub(rb'<Relationship [^>]*Target="(header|footer)\d*\.xml"[^>]*/>', b"", data)
            if item.filename == "[Content_Types].xml":
                data = re.sub(rb'<Override [^>]*PartName="/word/(header|footer)\d*\.xml"[^>]*/>', b"", data)
            o.writestr(item, data)
import win32com.client as w  # noqa: E402

app = w.Dispatch("Word.Application")
app.Visible = False
pdf = os.path.abspath(out)[:-5] + ".pdf"
try:
    d = app.Documents.Open(os.path.abspath(out), ReadOnly=True)
    d.ExportAsFixedFormat(pdf, 17)
    d.Close(False)
finally:
    app.Quit()
import fitz  # noqa: E402

pitch = int(re.search(r'w:linePitch="(\d+)"', sect).group(1)) / 20.0
top = int(re.search(r'w:top="(-?\d+)"', re.search(r"<w:pgMar [^>]*/>", sect).group(0)).group(1)) / 20.0
print("host %s | grid pitch %.2fpt | top margin %.2fpt | font %s" % (os.path.basename(src), pitch, top, FONT))
pd = fitz.open(pdf)
spans = [s for pno in range(1) for b in pd[pno].get_text("dict")["blocks"] for l in b.get("lines", []) for s in l["spans"] if s["text"].strip()]
spans.sort(key=lambda s: s["origin"][1])
print("arm  size   baseline   grid k  block(top..bottom)   em top   em box centred?  top-aligned?")
prev_bottom = top
for s in spans:
    size = s["size"]
    base = s["origin"][1]
    # the em box: ascent 0.8594 for MS fonts is the hhea ascent; use the PDF's own
    # ascender via the span bbox top instead, which is the ink, so report both
    need = size * 1.0
    n = max(1, int((need + 1e-6) // pitch) + (0 if abs(need % pitch) < 1e-6 else 1))
    block_top = prev_bottom
    block_bot = block_top + n * pitch
    centred_em_top = block_top + (n * pitch - size) / 2.0
    print("  %-4s %5.1f  %8.2f  n=%d   %7.2f..%7.2f   %7.2f   centred %+.2f   top %+.2f"
          % (s["text"][:3], size, base, n, block_top, block_bot,
             base - 0.8594 * size, (base - 0.8594 * size) - centred_em_top,
             (base - 0.8594 * size) - block_top))
    prev_bottom = block_bot
