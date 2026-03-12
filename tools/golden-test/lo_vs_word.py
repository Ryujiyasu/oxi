#!/usr/bin/env python3
"""
Pixel-level comparison: LibreOffice PDF vs Word PDF
Uses the same test documents and Word PNGs from pixel_compare.py.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import cv2
import fitz  # PyMuPDF
import numpy as np
from skimage.metrics import structural_similarity as ssim

WEB_ROOT = Path(__file__).resolve().parent.parent.parent
SOFFICE = Path("C:/Program Files/LibreOffice/program/soffice.exe")
DPI = 150
TIMEOUT_SEC = 60
OUTPUT_DIR = Path(__file__).resolve().parent / "pixel_output"
WORD_DIR = OUTPUT_DIR / "word"  # reuse Word PNGs from pixel_compare.py


def pdf_to_png(pdf_path: Path, out_png: Path, dpi: int = DPI) -> bool:
    try:
        doc = fitz.open(str(pdf_path))
        if len(doc) == 0:
            return False
        page = doc[0]
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        pix.save(str(out_png))
        doc.close()
        return True
    except Exception as e:
        print(f"    PDF->PNG error: {e}")
        return False


def render_lo_pdf(docx_path: Path, out_png: Path) -> dict:
    start = time.time()
    tmp = tempfile.mkdtemp(prefix="lo_pdf_")
    try:
        result = subprocess.run(
            [str(SOFFICE), "--headless", "--convert-to", "pdf",
             "--outdir", tmp, str(docx_path)],
            capture_output=True, text=True, timeout=TIMEOUT_SEC,
        )
        elapsed = int((time.time() - start) * 1000)

        pdf_path = Path(tmp) / (docx_path.stem + ".pdf")
        if not pdf_path.exists():
            return {"success": False, "error": "No PDF output", "time_ms": elapsed}

        if not pdf_to_png(pdf_path, out_png):
            return {"success": False, "error": "PDF->PNG failed", "time_ms": elapsed}

        return {"success": True, "error": None, "time_ms": elapsed}
    except subprocess.TimeoutExpired:
        return {"success": False, "error": "Timeout", "time_ms": int((time.time() - start) * 1000)}
    except Exception as e:
        return {"success": False, "error": str(e)[:200], "time_ms": int((time.time() - start) * 1000)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def imread_unicode(path: Path):
    """cv2.imread that supports Unicode paths."""
    try:
        buf = np.fromfile(str(path), dtype=np.uint8)
        return cv2.imdecode(buf, cv2.IMREAD_COLOR)
    except Exception:
        return None


def compare_images(img1_path: Path, img2_path: Path) -> dict:
    try:
        img1 = imread_unicode(img1_path)
        img2 = imread_unicode(img2_path)
        if img1 is None or img2 is None:
            return {"ssim": None, "error": "Failed to load image(s)"}
        h = min(img1.shape[0], img2.shape[0])
        w = min(img1.shape[1], img2.shape[1])
        img1 = cv2.resize(img1, (w, h))
        img2 = cv2.resize(img2, (w, h))
        gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
        gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        score, _ = ssim(gray1, gray2, full=True)
        return {"ssim": round(score, 4), "error": None}
    except Exception as e:
        return {"ssim": None, "error": str(e)[:200]}


def create_side_by_side(img_paths: dict, out_path: Path):
    imgs = []
    for label, path in img_paths.items():
        if path and Path(path).exists():
            img = imread_unicode(Path(path))
            if img is not None:
                h, w = img.shape[:2]
                labeled = np.zeros((h + 40, w, 3), dtype=np.uint8)
                labeled[:] = (255, 255, 255)
                cv2.putText(labeled, label, (10, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 0), 2)
                labeled[40:40 + h, :w] = img
                imgs.append(labeled)
    if len(imgs) < 2:
        return
    max_h = max(img.shape[0] for img in imgs)
    resized = []
    for img in imgs:
        h, w = img.shape[:2]
        if h < max_h:
            padded = np.ones((max_h, w, 3), dtype=np.uint8) * 255
            padded[:h, :w] = img
            resized.append(padded)
        else:
            resized.append(img)
    combined = np.hstack(resized)
    _, buf = cv2.imencode('.png', combined)
    buf.tofile(str(out_path))


def main():
    fixtures_dir = WEB_ROOT / "tests" / "fixtures"
    golden_docx_dir = Path(__file__).resolve().parent / "documents" / "docx"

    test_files = []
    for f in sorted(fixtures_dir.glob("*.docx")):
        test_files.append(f)
    if golden_docx_dir.exists():
        golden_files = sorted(golden_docx_dir.glob("*.docx"))
        test_files.extend(golden_files)

    if not test_files:
        print("No docx files found")
        sys.exit(1)

    total = len(test_files)

    print("=" * 70)
    print("  Pixel-Level Comparison: LibreOffice PDF vs Word PDF")
    print("=" * 70)
    print(f"  Files: {total}")
    print(f"  DPI: {DPI}")
    print(f"  LibreOffice: {SOFFICE}")
    print("=" * 70)
    print()

    lo_dir = OUTPUT_DIR / "libreoffice"
    lo_dir.mkdir(parents=True, exist_ok=True)
    lo_compare_dir = OUTPUT_DIR / "lo_comparison"
    lo_compare_dir.mkdir(parents=True, exist_ok=True)

    # Also load Oxi PNGs for 3-way comparison
    oxi_dir = OUTPUT_DIR / "oxi"

    # Load existing Oxi vs Word report for side-by-side summary
    oxi_report_path = OUTPUT_DIR / "pixel_comparison_report.json"
    oxi_ssims = {}
    if oxi_report_path.exists():
        oxi_report = json.loads(oxi_report_path.read_text())
        for comp in oxi_report.get("comparisons", []):
            s = comp.get("oxi_vs_word", {}).get("ssim")
            if s is not None:
                oxi_ssims[comp["filename"]] = s

    comparisons = []

    for i, docx_path in enumerate(test_files):
        fname = docx_path.name
        stem = docx_path.stem
        print(f"[{i + 1:>3}/{total:>3}] {fname[:50]}", end="", flush=True)

        lo_png = lo_dir / f"{stem}.png"
        lo_res = render_lo_pdf(docx_path, lo_png)
        lo_status = "OK" if lo_res["success"] else "FAIL"
        print(f"  LO:{lo_status}({lo_res['time_ms']}ms)", end="", flush=True)

        word_png = WORD_DIR / f"{stem}.png"
        oxi_png = oxi_dir / f"{stem}.png"

        comp = {"filename": fname}

        if lo_png.exists() and word_png.exists():
            comp["lo_vs_word"] = compare_images(lo_png, word_png)
        else:
            comp["lo_vs_word"] = {"ssim": None, "error": "Missing image(s)"}

        comp["oxi_vs_word_ssim"] = oxi_ssims.get(fname)
        comparisons.append(comp)

        # Create 3-way side-by-side: Oxi | LO | Word
        side_images = {}
        if oxi_png.exists():
            side_images["Oxi-PDF"] = str(oxi_png)
        if lo_png.exists():
            side_images["LO-PDF"] = str(lo_png)
        if word_png.exists():
            side_images["Word-PDF"] = str(word_png)
        if len(side_images) >= 2:
            create_side_by_side(side_images, lo_compare_dir / f"{stem}_compare.png")

        lo_ssim = comp["lo_vs_word"].get("ssim", "-")
        oxi_s = comp["oxi_vs_word_ssim"]
        oxi_str = f"{oxi_s:.4f}" if oxi_s else "-"
        print(f"  LO-SSIM: {lo_ssim}  Oxi-SSIM: {oxi_str}", flush=True)

    # Summary
    print("\n" + "=" * 70)
    print("  SSIM SUMMARY: LibreOffice vs Word  |  Oxi vs Word")
    print("=" * 70)

    lo_vals = [c["lo_vs_word"]["ssim"] for c in comparisons if c["lo_vs_word"].get("ssim") is not None]
    oxi_vals = [c["oxi_vs_word_ssim"] for c in comparisons if c.get("oxi_vs_word_ssim") is not None]

    if lo_vals:
        print(f"  LO  Average: {sum(lo_vals)/len(lo_vals):.4f}   Best: {max(lo_vals):.4f}   Worst: {min(lo_vals):.4f}   Files: {len(lo_vals)}/{total}")
    if oxi_vals:
        print(f"  Oxi Average: {sum(oxi_vals)/len(oxi_vals):.4f}   Best: {max(oxi_vals):.4f}   Worst: {min(oxi_vals):.4f}   Files: {len(oxi_vals)}/{total}")

    # Per-file winner
    print("\n  Per-file comparison:")
    oxi_wins = 0
    lo_wins = 0
    for c in comparisons:
        lo_s = c["lo_vs_word"].get("ssim")
        oxi_s = c.get("oxi_vs_word_ssim")
        if lo_s is not None and oxi_s is not None:
            winner = "Oxi" if oxi_s >= lo_s else "LO"
            if oxi_s >= lo_s:
                oxi_wins += 1
            else:
                lo_wins += 1
            print(f"    {c['filename'][:45]:45}  LO={lo_s:.4f}  Oxi={oxi_s:.4f}  {winner}")

    print(f"\n  Oxi wins: {oxi_wins}   LO wins: {lo_wins}")
    print("=" * 70)

    report = {
        "comparisons": comparisons,
        "summary": {
            "lo_avg": sum(lo_vals) / len(lo_vals) if lo_vals else None,
            "oxi_avg": sum(oxi_vals) / len(oxi_vals) if oxi_vals else None,
            "oxi_wins": oxi_wins,
            "lo_wins": lo_wins,
        },
    }
    report_path = OUTPUT_DIR / "lo_vs_word_report.json"
    report_path.write_text(json.dumps(report, indent=2, default=str))
    print(f"\nReport: {report_path}")
    print(f"3-way images: {lo_compare_dir}")


if __name__ == "__main__":
    main()
