"""Render V1-V6 minimal repros via Word, then measure glyph y from PNG.

This validates whether Word's PNG-rendered cell paragraph spacing matches the
COM Information(6) cursor advance (=13pt for V3-V6) or differs (suggesting
Information(6) is not glyph top).

Output: pipeline_data/b5f706_l7_variants_word_png_measured.json
"""
from __future__ import annotations

import glob
import json
import os
import sys
import time
import subprocess
from pathlib import Path

import numpy as np
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
REPRO_DIR = REPO / "tools" / "golden-test" / "repros" / "grid_snap"
WORD_PNG_DIR = REPO / "pipeline_data" / "word_png_b5f706_variants"
DPI = 150
OUT = REPO / "pipeline_data" / "b5f706_l7_variants_word_png_measured.json"


_RENDER_SCRIPT = r'''
import sys, os
docx_path, out_dir, dpi = sys.argv[1], sys.argv[2], int(sys.argv[3])
import win32com.client
import pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
try:
    doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
    try:
        page_count = doc.ComputeStatistics(2)
        for page_num in range(1, page_count + 1):
            pdf_path = os.path.join(out_dir, f"page_{page_num:04d}.pdf")
            png_path = os.path.join(out_dir, f"page_{page_num:04d}.png")
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=17, OpenAfterExport=False,
                OptimizeFor=0, Range=3, From=page_num, To=page_num,
            )
            import fitz
            d = fitz.open(pdf_path)
            zoom = dpi / 72
            pix = d[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            pix.save(png_path)
            d.close()
            os.unlink(pdf_path)
    finally:
        doc.Close(SaveChanges=False)
finally:
    word.Quit()
    pythoncom.CoUninitialize()
'''


def render_one(docx_path: Path, out_dir: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    existing = sorted(out_dir.glob("page_*.png"))
    if existing:
        return existing
    try:
        result = subprocess.run(
            [sys.executable, "-c", _RENDER_SCRIPT,
             str(docx_path.resolve()), str(out_dir), str(DPI)],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=30,
        )
        if result.returncode != 0:
            print(f"  Word error: {result.stderr[:200]}")
            return []
        return sorted(out_dir.glob("page_*.png"))
    except Exception as e:
        print(f"  Render exception: {e}")
        return []


def detect_text_lines(png_path: Path, dpi: int = 150) -> list[float]:
    """Detect glyph baseline y positions by row-projection of dark pixels.

    Returns list of y_pt where text appears (centroid of each non-blank row).
    """
    img = np.array(Image.open(png_path).convert('L'))
    # text mask: pixels darker than 200
    dark = img < 200
    row_count = dark.sum(axis=1)
    # threshold: row has > 5 dark pixels
    is_text_row = row_count > 5

    # Find runs of consecutive text rows -> each run = one line of text
    lines = []
    in_run = False
    run_start = 0
    for r, t in enumerate(is_text_row):
        if t and not in_run:
            in_run = True
            run_start = r
        elif not t and in_run:
            in_run = False
            run_end = r
            # baseline approximation = bottom of run minus typical descender
            line_top = run_start
            line_bottom = run_end
            line_h_px = line_bottom - line_top
            # use top of glyph cluster as line top
            lines.append(line_top * 72.0 / dpi)
    if in_run:
        lines.append(run_start * 72.0 / dpi)
    return lines


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    docx_files = sorted(REPRO_DIR.glob("b5f706_V*.docx"))
    if not docx_files:
        print(f"No V*.docx in {REPRO_DIR}")
        return

    results = []
    for docx in docx_files:
        label = docx.stem
        out_dir = WORD_PNG_DIR / label
        print(f"=== {label} ===")
        pngs = render_one(docx, out_dir)
        if not pngs:
            print(f"  Failed to render")
            continue
        # measure first PNG
        png = pngs[0]
        lines = detect_text_lines(png, DPI)
        deltas = [round(lines[i] - lines[i-1], 2) for i in range(1, len(lines))]
        # filter outliers (e.g. table border lines, header)
        small_deltas = [d for d in deltas if 5 < d < 30]
        avg_delta = round(sum(small_deltas) / len(small_deltas), 3) if small_deltas else None
        print(f"  PNG: {png.name}, lines: {len(lines)}, glyph y: {[round(y,1) for y in lines[:8]]}")
        print(f"  deltas: {deltas[:7]}")
        print(f"  avg_glyph_dy (5..30pt window): {avg_delta}")
        results.append({
            "label": label,
            "png": str(png),
            "n_lines": len(lines),
            "glyph_y": [round(y, 2) for y in lines],
            "glyph_deltas": deltas,
            "avg_glyph_dy_filtered": avg_delta,
        })

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"results": results}, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
