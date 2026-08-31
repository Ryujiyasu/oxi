# -*- coding: utf-8 -*-
"""Where does the first line's INK start? Word vs Oxi, with no shared convention.

Every earlier comparison of "where the text sits" had to pick a frame -- Word's
PDF baseline, Word's Information(6), Oxi's element y, Oxi's y+text_y_off, the
glyph cell top -- and each pairing needs a metric neither side agrees on. Ink
needs none: rasterise both at the same DPI and find the first row that has any
dark pixel. S1097 says Oxi's 0.5pt top-margin round is compensating ~0.25pt of
vertical error somewhere; this measures that number directly.

    INKTOP_DPI=600 python tools/metrics/_pb_inktop.py <docx> [<docx> ...]
"""
import os, subprocess, sys, glob
import fitz
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Which renderer? The SSIM gate uses DWrite; GDI is the pagination/fallback
# leg and places its glyphs with GDI's OWN integer tmAscent, so the two do
# NOT agree on where a baseline sits. Say which one you are asking.
ENGINE = os.environ.get("INKTOP_ENGINE", "gdi")
REND = os.path.abspath("tools/oxi-%s-renderer/target/release/oxi-%s-renderer.exe"
                       % (ENGINE, ENGINE))
DPI = int(os.environ.get("INKTOP_DPI", "300"))
THRESH = 160          # a pixel darker than this counts as ink


def first_ink_row(pix):
    w, h, n = pix.width, pix.height, pix.n
    data = pix.samples
    for y in range(h):
        base = y * pix.stride
        for x in range(0, w * n, n):
            if data[base + x] < THRESH:
                return y
    return None


def word_ink_top(docx):
    import win32com.client
    pdf = docx[:-5] + ".pdf"
    if not os.path.exists(pdf):
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
            d.Close(False)
        finally:
            app.Quit()
    doc = fitz.open(pdf)
    pix = doc[0].get_pixmap(dpi=DPI, colorspace=fitz.csGRAY)
    row = first_ink_row(pix)
    doc.close()
    return row


def oxi_ink_top(docx):
    prefix = docx[:-5] + "_ink_" + ENGINE
    for old in glob.glob(prefix + "*.png"):
        os.remove(old)
    subprocess.run([REND, docx, prefix, str(DPI)], capture_output=True)
    pngs = sorted(glob.glob(prefix + "*.png"))
    if not pngs:
        return None
    pix = fitz.Pixmap(pngs[0])
    if pix.n > 1:
        pix = fitz.Pixmap(fitz.csGRAY, pix)
    return first_ink_row(pix)


print("  %-26s %10s %10s %10s %10s" % ("arm", "word_px", "oxi_px", "d_px", "d_pt"))
for docx in sys.argv[1:]:
    w = word_ink_top(docx)
    o = oxi_ink_top(docx)
    if w is None or o is None:
        print("  %-26s  (word %s oxi %s)" % (os.path.basename(docx)[:-5], w, o))
        continue
    print("  %-26s %10d %10d %10d %10.3f"
          % (os.path.basename(docx)[:-5], w, o, o - w, (o - w) * 72.0 / DPI))
