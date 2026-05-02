"""Measure 1ec1's actual □3 paragraph via my PDF→fitz pipeline.

If user's PNG measurement gives 48pt and my PDF measurement gives different,
the +5.5pt synthetic vs 1ec1 gap is a measurement-method artifact, not a
structural property difference.
"""
import os
import sys
import time
import json
import pythoncom
import win32com.client
import fitz
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_actual_pdf_measure")
os.makedirs(OUT_DIR, exist_ok=True)
PDF = os.path.join(OUT_DIR, "1ec1.pdf")
RESULT = os.path.join(OUT_DIR, "results.json")


def render_pdf(word, docx_path, pdf_path):
    last = None
    for attempt in range(5):
        try:
            doc = word.Documents.Open(docx_path, ReadOnly=True)
            time.sleep(0.5)
            doc.SaveAs2(pdf_path, FileFormat=17)
            doc.Close(SaveChanges=False)
            return True
        except Exception as e:
            last = e
            time.sleep(1.0 + attempt * 0.5)
    return False


def find_box_pages(pdf_path):
    """Find pages that contain □ char and identify pixel x positions of all dark text."""
    d = fitz.open(pdf_path)
    pages_with_box = []
    for page_idx in range(min(5, d.page_count)):
        page = d[page_idx]
        text = page.get_text()
        if "□" in text:
            pages_with_box.append(page_idx)
            # Find precise positions of □ via search
            instances = page.search_for("□")
            zoom = 4.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            w, h, n = pix.width, pix.height, pix.n
            s = pix.samples
            # For each □ instance, get its rect in pt and convert to pixel range
            print(f"\nPage {page_idx + 1}: found {len(instances)} '□' chars")
            for i, rect in enumerate(instances):
                # rect: PyMuPDF Rect in pt
                left_pt = rect.x0
                right_pt = rect.x1
                top_pt = rect.y0
                bottom_pt = rect.y1
                left_px = int(left_pt * zoom)
                right_px = int(right_pt * zoom)
                top_px = int(top_pt * zoom)
                bottom_px = int(bottom_pt * zoom)
                # Find leftmost dark pixel in this rect's y range
                leftmost = None
                for py in range(max(0, top_px), min(h, bottom_px)):
                    for px in range(max(0, left_px - 20), min(w, right_px + 5)):
                        off = (py * w + px) * n
                        r, g, b = s[off], s[off+1], s[off+2]
                        if r < 200 and g < 200 and b < 200:
                            if leftmost is None or px < leftmost:
                                leftmost = px
                            break
                left_pt_actual = leftmost / zoom if leftmost else None
                print(f"  □ #{i+1}: search rect L={left_pt:.2f}pt R={right_pt:.2f}pt T={top_pt:.2f} B={bottom_pt:.2f}")
                print(f"    leftmost dark pixel: x={leftmost} ({left_pt_actual:.2f}pt)" if leftmost else "    no dark pixel found")
                if i >= 3: break  # only first few
    d.close()
    return pages_with_box


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    try:
        ok = render_pdf(word, DOCX, PDF)
        if not ok:
            print("PDF render failed")
            return
        find_box_pages(PDF)
    finally:
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
