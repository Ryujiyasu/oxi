# -*- coding: utf-8 -*-
"""Compare Oxi chart renders against the Word PDFs (150dpi).

Usage: python tools/metrics/chart_ssim.py <doc_id> <oxi_png> [page_index]
e.g.  python tools/metrics/chart_ssim.py chart1 scratchpad/chart_step5/render_chart1/c1_s1.png
      python tools/metrics/chart_ssim.py chart2 scratchpad/chart_step5/render_chart2/c2_s1.png
      python tools/metrics/chart_ssim.py chart2b scratchpad/chart_step5/render_chart2b/c2b_s1.png
      python tools/metrics/chart_ssim.py chart3 scratchpad/chart_step5/render_chart3/c3_s1.png
      python tools/metrics/chart_ssim.py chart_pie2 scratchpad/pie_render/chart_pie2_s3.png 2
"""
import sys
import numpy as np
import fitz
from PIL import Image

DOCS = {
    "chart1": r"pipeline_data\pptx_probes\chart1\chart1.pdf",
    "chart2": r"pipeline_data\pptx_probes\chart2\chart2.pdf",
    "chart2b": r"pipeline_data\pptx_probes\chart2b\chart2b.pdf",
    "chart3": r"pipeline_data\pptx_probes\chart3\chart3.pdf",
    "chart_legend": r"pipeline_data\pptx_probes\chart_legend\chart_legend.pdf",
    "chart_legend3": r"pipeline_data\pptx_probes\chart_legend3\chart_legend3.pdf",
    "chart_pie": r"pipeline_data\pptx_probes\chart_pie\chart_pie.pdf",
    "chart_pie3": r"pipeline_data\pptx_probes\chart_pie3\chart_pie3.pdf",
    "chart_pie2": r"pipeline_data\pptx_probes\chart_pie2\chart_pie2.pdf",
    "chart_stacked": r"pipeline_data\pptx_probes\chart_stacked\chart_stacked.pdf",
    "chart_stacked100": r"pipeline_data\pptx_probes\chart_stacked100\chart_stacked100.pdf",
    "chart_title": r"pipeline_data\pptx_probes\chart_title\chart_title.pdf",
    "chart_title2": r"pipeline_data\pptx_probes\chart_title2\chart_title2.pdf",
    "chart_line": r"pipeline_data\pptx_probes\chart_line\chart_line.pdf",
    "chart_line2": r"pipeline_data\pptx_probes\chart_line2\chart_line2.pdf",
    "chart_line3": r"pipeline_data\pptx_probes\chart_line3\chart_line3.pdf",
}
DPI = 150


def ink_bbox(gray: np.ndarray):
    mask = gray < 200
    if not mask.any():
        return None
    ys, xs = np.nonzero(mask)
    return (int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max()))


def main() -> None:
    doc_id = sys.argv[1]
    oxi_png = sys.argv[2]
    word_pdf = DOCS[doc_id]

    doc = fitz.open(word_pdf)
    page_idx = int(sys.argv[3]) if len(sys.argv) > 3 else 0
    pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(DPI / 72, DPI / 72))
    word = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, pix.n
    )
    word_rgb = word[..., :3] if pix.n >= 3 else np.repeat(word[..., None], 3, axis=2)
    word_gray = word_rgb.mean(axis=2)
    print("Word:", word_pdf, pix.width, "x", pix.height)

    oxi = np.array(Image.open(oxi_png).convert("RGB"))
    print("Oxi :", oxi_png, oxi.shape[1], "x", oxi.shape[0])

    h = min(word_gray.shape[0], oxi.shape[0])
    w = min(word_gray.shape[1], oxi.shape[1])
    wg = word_gray[:h, :w]
    og = oxi[:h, :w].mean(axis=2)
    wb = word_rgb[:h, :w]
    ob = oxi[:h, :w]

    print("Word ink bbox:", ink_bbox(wg))
    print("Oxi  ink bbox:", ink_bbox(og))

    from skimage.metrics import structural_similarity as ssim
    print("SSIM (grayscale, common area): %.4f" % ssim(wg, og, data_range=255))

    # Ink pixel counts + dominant colours.
    from collections import Counter
    for name, arr in (("Word", wb), ("Oxi", ob)):
        nw = (arr < 200).any(axis=2)
        print("%s non-white px: %d" % (name, int(nw.sum())))
        if nw.any():
            cols = arr[nw]
            keyed = (
                cols[:, 0].astype(int) * 65536
                + cols[:, 1].astype(int) * 256
                + cols[:, 2].astype(int)
            )
            top = Counter(keyed).most_common(5)
            print("%s top colours:" % name)
            for k, c in top:
                print("   #%02X%02X%02X  x%d" % ((k >> 16) & 255, (k >> 8) & 255, k & 255, c))


if __name__ == "__main__":
    main()
