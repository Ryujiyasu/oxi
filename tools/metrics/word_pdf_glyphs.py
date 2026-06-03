"""Render-truth via PDF (robust for tables, where EMF export fails).
Word ExportAsFixedFormat -> PDF; PyMuPDF extracts per-glyph (x,y,char) in points.
PDF uses Word's layout engine = same positions as the screenshot/EMF = render-truth.
Usage: python word_pdf_glyphs.py <docx> <out.json>   (exports <docx>.pdf alongside)
cp932-safe: no Japanese in code; per-glyph text written to JSON (UTF-8)."""
import sys, os, json


def export_pdf(docx):
    import win32com.client, pythoncom
    docx = os.path.abspath(docx)
    pdf = os.path.splitext(docx)[0] + "_rt.pdf"
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(docx, ReadOnly=True)
        doc.ExportAsFixedFormat(pdf, 17)  # wdExportFormatPDF=17
        doc.Close(False)
    finally:
        word.Quit()
    return pdf


def extract(pdf, out):
    import fitz
    d = fitz.open(pdf)
    pages = []
    for pno in range(d.page_count):
        page = d[pno]
        rd = page.get_text("rawdict")
        glyphs = []
        for blk in rd.get("blocks", []):
            for line in blk.get("lines", []):
                for span in line.get("spans", []):
                    size = span.get("size", 0)
                    for ch in span.get("chars", []):
                        c = ch["c"]
                        if not c.strip():
                            continue
                        ox, oy = ch["origin"]  # baseline-left, in points (72dpi)
                        glyphs.append({"char": c, "x": round(ox, 2), "y": round(oy, 2),
                                       "fs": round(size, 1)})
        pages.append({"glyphs": glyphs, "w": round(page.rect.width, 1),
                      "h": round(page.rect.height, 1)})
    json.dump({"pages": pages}, open(out, "w", encoding="utf-8"), ensure_ascii=False)
    print("pages=%d  per-page glyphs: %s" % (len(pages), [len(p["glyphs"]) for p in pages]))


if __name__ == "__main__":
    docx, out = sys.argv[1], sys.argv[2]
    pdf = export_pdf(docx)
    print("PDF:", pdf)
    extract(pdf, out)
    print("wrote", out)
