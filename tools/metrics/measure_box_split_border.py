"""Measure Word's close-border y on repros + d77a.

For each docx:
1. Open in Word, export each page to PNG
2. Find horizontal lines on page with row-split (page N where table splits)
3. Measure:
   - last_content_bottom (via COM Info(6) + line_height of last para's last line)
   - close_border_y (from PNG scan, bottom-most long horizontal line in table area)
   - tcMar bottom (from XML)
   - page_bottom = page.height - bottom_margin
4. Derive formula: close_border_y = f(last_content_bottom, pad_b, page_bottom)
"""
import win32com.client
import os
import sys
import json
from pathlib import Path
from PIL import Image
import numpy as np

import glob as _glob
REPROS = sorted(_glob.glob(r"C:\Users\ryuji\oxi-main\tools\metrics\box_split_repro\repro_*.docx"))
D77A = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"

OUT_DIR = Path(r"C:\Users\ryuji\oxi-main\tools\metrics\box_split_repro\word_png")
OUT_DIR.mkdir(exist_ok=True, parents=True)


def export_word_pages_to_png(docx_path: str, out_prefix: str, pages_to_export: list[int] | None = None) -> list[str]:
    """Open docx in Word, export each page as PNG via CopyAsPicture + metafile.
    For simplicity, use Word's SaveAs PDF and then convert or use InlineShape approach.
    Actually fastest: use Word's built-in screen render via SaveAs XPS or PDF + pdfplumber.
    Alternative: use word.ActiveDocument.InlineShapes... (limited).

    Simpler: use the existing pipeline's Word renderer.
    """
    # Just return paths as guidance
    return []


def measure_com(docx_path: str) -> dict:
    """Get paragraph-level data from Word COM."""
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        total = doc.Paragraphs.Count
        paragraphs = []
        for i in range(1, total + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            try:
                page_start = rng.Information(3)
                y_start = rng.Information(6)
            except Exception:
                page_start = -1; y_start = -1.0
            try:
                safe_end = max(rng.Start, rng.End - 1)
                end_rng = doc.Range(safe_end, safe_end)
                page_end = end_rng.Information(3)
                y_end = end_rng.Information(6)
            except Exception:
                page_end = -1; y_end = -1.0
            try:
                x_start = rng.Information(5)
            except Exception:
                x_start = -1.0
            text = rng.Text[:50].replace("\r", " ").replace("\n", " ")
            paragraphs.append({
                "idx": i, "page_start": page_start, "page_end": page_end,
                "y_start": round(y_start, 2), "y_end": round(y_end, 2),
                "x_start": round(x_start, 2), "text": text,
            })

        sec = doc.Sections(1)
        page_h = sec.PageSetup.PageHeight
        page_w = sec.PageSetup.PageWidth
        top_m = sec.PageSetup.TopMargin
        bot_m = sec.PageSetup.BottomMargin
        left_m = sec.PageSetup.LeftMargin
        right_m = sec.PageSetup.RightMargin

        doc.Close(SaveChanges=False)
        return {
            "page_h": page_h, "page_w": page_w,
            "top_m": top_m, "bot_m": bot_m, "left_m": left_m, "right_m": right_m,
            "paragraphs": paragraphs,
        }
    finally:
        word.Quit()


def measure_word_png(png_path: str, page_h_pt: float = 841.89) -> dict:
    """Find horizontal lines in a Word PNG rendering."""
    img = np.array(Image.open(png_path).convert("L"))
    h, w = img.shape
    pt_per_px = page_h_pt / h
    lines = []
    for y in range(h):
        n_dark = int(np.sum(img[y] < 100))
        if n_dark > 400:
            lines.append({"y_px": y, "y_pt": round(y * pt_per_px, 2), "dark_px": n_dark})
    # Merge contiguous
    merged = []
    cur = None
    for l in lines:
        if cur is None or l["y_px"] - cur["y_px_max"] > 2:
            if cur: merged.append(cur)
            cur = {"y_px_min": l["y_px"], "y_px_max": l["y_px"], "y_pt_min": l["y_pt"], "y_pt_max": l["y_pt"], "dark_px": l["dark_px"]}
        else:
            cur["y_px_max"] = l["y_px"]; cur["y_pt_max"] = l["y_pt"]
            cur["dark_px"] = max(cur["dark_px"], l["dark_px"])
    if cur: merged.append(cur)
    return {"image_size": {"w": w, "h": h}, "pt_per_px": round(pt_per_px, 4), "h_lines": merged}


def main():
    all_data = {}
    for docx in REPROS + [D77A]:
        name = Path(docx).stem
        print(f"\n=== {name} ===", flush=True)
        com = measure_com(docx)
        all_data[name] = {"docx": docx, "com": com}

        # For repros, also scan Word PNG via existing pipeline word_png dir
        # For d77a, use pipeline_data/word_png
        word_png_dir = Path(r"C:\Users\ryuji\oxi-main\pipeline_data\word_png") / name
        if word_png_dir.exists():
            page_data = []
            for png in sorted(word_png_dir.glob("page_*.png")):
                pd = measure_word_png(str(png), com["page_h"])
                pd["page_num"] = int(png.stem.split("_")[-1])
                page_data.append(pd)
            all_data[name]["word_png"] = page_data
            print(f"  Found Word PNGs: {len(page_data)}")
        else:
            print(f"  No Word PNG dir at {word_png_dir}")

    out = Path(r"C:\Users\ryuji\oxi-main\pipeline_data\box_split_measurements.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
