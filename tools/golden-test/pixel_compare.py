#!/usr/bin/env python3
"""
Pixel-level comparison: Oxi PDF vs Word PDF (primary)
Renders each .docx file with oxi-cli (docx-to-pdf) and Word COM, then compares PNG output using SSIM.

Requirements: PyMuPDF, scikit-image, opencv-python, numpy, pillow, pywin32
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
from PIL import Image
from skimage.metrics import structural_similarity as ssim

# --- Configuration ---
WEB_ROOT = Path(__file__).resolve().parent.parent.parent  # oxi-1/
OXI_CLI = WEB_ROOT / "target" / "release" / "oxi.exe"
DPI = 150  # render resolution
TIMEOUT_SEC = 60
OUTPUT_DIR = Path(__file__).resolve().parent / "pixel_output"

# --- Helpers ---

def pdf_to_png(pdf_path: Path, out_png: Path, dpi: int = DPI) -> bool:
    """Convert first page of PDF to PNG using PyMuPDF."""
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


def render_oxi_pdf(docx_path: Path, out_png: Path) -> dict:
    """Render docx via oxi-cli docx-to-pdf -> PDF -> PNG."""
    start = time.time()
    tmp = tempfile.mkdtemp(prefix="oxi_pdf_")

    try:
        pdf_path = Path(tmp) / (docx_path.stem + ".pdf")
        result = subprocess.run(
            [str(OXI_CLI), "docx-to-pdf", str(docx_path), str(pdf_path)],
            capture_output=True, text=True, timeout=TIMEOUT_SEC,
        )
        elapsed = int((time.time() - start) * 1000)

        if not pdf_path.exists():
            stderr = result.stderr[:200] if result.stderr else "No PDF output"
            return {"success": False, "error": stderr, "time_ms": elapsed}

        if not pdf_to_png(pdf_path, out_png):
            return {"success": False, "error": "PDF->PNG failed", "time_ms": elapsed}

        return {"success": True, "error": None, "time_ms": elapsed}
    except subprocess.TimeoutExpired:
        return {"success": False, "error": "Timeout", "time_ms": int((time.time() - start) * 1000)}
    except Exception as e:
        return {"success": False, "error": str(e)[:200], "time_ms": int((time.time() - start) * 1000)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def render_word(docx_path: Path, out_png: Path) -> dict:
    """Render docx via Microsoft Word COM -> PDF -> PNG."""
    start = time.time()
    tmp = tempfile.mkdtemp(prefix="word_pixel_")

    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        pdf_path = Path(tmp) / (docx_path.stem + ".pdf")

        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        doc.SaveAs2(str(pdf_path.resolve()), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close(SaveChanges=0)
        word.Quit()

        elapsed = int((time.time() - start) * 1000)

        if not pdf_path.exists():
            return {"success": False, "error": "No PDF output", "time_ms": elapsed}

        if not pdf_to_png(pdf_path, out_png):
            return {"success": False, "error": "PDF->PNG failed", "time_ms": elapsed}

        return {"success": True, "error": None, "time_ms": elapsed}
    except Exception as e:
        try:
            subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"],
                           capture_output=True, timeout=5)
        except Exception:
            pass
        return {"success": False, "error": str(e)[:200], "time_ms": int((time.time() - start) * 1000)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def imread_unicode(path: Path):
    """cv2.imread that supports Unicode paths."""
    try:
        buf = np.fromfile(str(path), dtype=np.uint8)
        return cv2.imdecode(buf, cv2.IMREAD_COLOR)
    except Exception:
        return None


def compare_images(img1_path: Path, img2_path: Path) -> dict:
    """Compare two images using SSIM and pixel diff metrics."""
    try:
        img1 = imread_unicode(img1_path)
        img2 = imread_unicode(img2_path)

        if img1 is None or img2 is None:
            return {"ssim": None, "error": "Failed to load image(s)"}

        # Resize to same dimensions (use the smaller of the two)
        h1, w1 = img1.shape[:2]
        h2, w2 = img2.shape[:2]
        h = min(h1, h2)
        w = min(w1, w2)
        img1 = cv2.resize(img1, (w, h))
        img2 = cv2.resize(img2, (w, h))

        # Convert to grayscale for SSIM
        gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
        gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)

        # SSIM
        score, diff_map = ssim(gray1, gray2, full=True)

        # Pixel-level diff
        diff = cv2.absdiff(img1, img2)
        diff_pixels = np.count_nonzero(cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY))
        total_pixels = h * w
        diff_ratio = diff_pixels / total_pixels

        return {
            "ssim": round(score, 4),
            "diff_pixel_ratio": round(diff_ratio, 4),
            "size1": f"{w1}x{h1}",
            "size2": f"{w2}x{h2}",
            "error": None,
        }
    except Exception as e:
        return {"ssim": None, "error": str(e)[:200]}


def create_side_by_side(images: dict, out_path: Path, labels: list):
    """Create a side-by-side comparison image."""
    imgs = []
    for label in labels:
        path = images.get(label)
        if path and Path(path).exists():
            img = imread_unicode(Path(path))
            if img is not None:
                h, w = img.shape[:2]
                labeled = np.zeros((h + 40, w, 3), dtype=np.uint8)
                labeled[:] = (255, 255, 255)
                cv2.putText(labeled, label, (10, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 0), 2)
                labeled[40:40 + h, :w] = img
                imgs.append(labeled)
            else:
                imgs.append(np.zeros((400, 300, 3), dtype=np.uint8))
        else:
            placeholder = np.ones((400, 300, 3), dtype=np.uint8) * 200
            cv2.putText(placeholder, f"{label}: N/A", (10, 200), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 0), 1)
            imgs.append(placeholder)

    if not imgs:
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
    # Collect docx files to test
    fixtures_dir = WEB_ROOT / "tests" / "fixtures"
    golden_docx_dir = Path(__file__).resolve().parent / "documents" / "docx"

    test_files = []

    # Add fixture files
    for f in sorted(fixtures_dir.glob("*.docx")):
        test_files.append(f)

    # Add a sample of golden test docx files (first 10)
    if golden_docx_dir.exists():
        golden_files = sorted(golden_docx_dir.glob("*.docx"))[:10]
        test_files.extend(golden_files)

    if not test_files:
        print("No docx files found for testing")
        sys.exit(1)

    total = len(test_files)

    # Check available renderers
    has_oxi = OXI_CLI.exists()
    has_word = False
    try:
        import win32com.client
        has_word = True
    except ImportError:
        pass

    print("=" * 70)
    print("  Pixel-Level Comparison: Oxi PDF vs Word PDF")
    print("=" * 70)
    print(f"  Files: {total}")
    print(f"  DPI: {DPI}")
    print(f"  oxi-cli:     {'FOUND' if has_oxi else 'NOT FOUND'}")
    print(f"  Word:         {'FOUND (COM)' if has_word else 'NOT FOUND'}")
    print("=" * 70)
    print()

    if not has_oxi:
        print(f"ERROR: oxi-cli not found at {OXI_CLI}")
        print("Run: cargo build -p oxi-cli --release")
        sys.exit(1)

    # Prepare output directories
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for sub in ["oxi", "word", "comparison"]:
        (OUTPUT_DIR / sub).mkdir(parents=True, exist_ok=True)

    # Render all files
    oxi_results = {}
    word_results = {}

    for i, docx_path in enumerate(test_files):
        fname = docx_path.name
        stem = docx_path.stem
        print(f"[{i + 1:>3}/{total:>3}] {fname[:50]}", end="", flush=True)

        # Oxi PDF
        oxi_png = OUTPUT_DIR / "oxi" / f"{stem}.png"
        oxi_res = render_oxi_pdf(docx_path, oxi_png)
        oxi_results[fname] = {**oxi_res, "png": str(oxi_png) if oxi_res["success"] else None}
        oxi_status = "OK" if oxi_res["success"] else "FAIL"
        print(f"  Oxi:{oxi_status}({oxi_res['time_ms']}ms)", end="", flush=True)

        # Word PDF (primary target) — skip if cached image exists
        if has_word:
            word_png = OUTPUT_DIR / "word" / f"{stem}.png"
            if word_png.exists():
                word_results[fname] = {"success": True, "error": None, "time_ms": 0, "png": str(word_png)}
                print(f"  Word:CACHED", end="", flush=True)
            else:
                word_res = render_word(docx_path, word_png)
                word_results[fname] = {**word_res, "png": str(word_png) if word_res["success"] else None}
                word_status = "OK" if word_res["success"] else "FAIL"
                print(f"  Word:{word_status}({word_res['time_ms']}ms)", end="", flush=True)
                time.sleep(0.3)

        print(flush=True)

    # Compare images
    print("\n>>> Computing SSIM comparisons...\n", flush=True)
    comparisons = []

    for docx_path in test_files:
        fname = docx_path.name
        stem = docx_path.stem

        oxi_png = OUTPUT_DIR / "oxi" / f"{stem}.png"
        word_png = OUTPUT_DIR / "word" / f"{stem}.png"

        comp = {"filename": fname}

        # Oxi PDF vs Word PDF (PRIMARY comparison)
        if oxi_png.exists() and word_png.exists():
            comp["oxi_vs_word"] = compare_images(oxi_png, word_png)
        else:
            comp["oxi_vs_word"] = {"ssim": None, "error": "Missing image(s)"}

        comparisons.append(comp)

        # Create side-by-side
        side_images = {}
        if oxi_png.exists():
            side_images["Oxi-PDF"] = str(oxi_png)
        if word_png.exists():
            side_images["Word-PDF"] = str(word_png)
        if len(side_images) >= 2:
            create_side_by_side(side_images,
                                OUTPUT_DIR / "comparison" / f"{stem}_compare.png",
                                list(side_images.keys()))

        # Print SSIM
        oxi_word_ssim = comp["oxi_vs_word"].get("ssim", "-")
        print(f"  {fname[:50]:50}  SSIM: {oxi_word_ssim}", flush=True)

    # Summary
    print("\n" + "=" * 70)
    print("  SSIM SUMMARY: Oxi PDF vs Word PDF (1.0 = identical)")
    print("=" * 70)

    vals = [c["oxi_vs_word"]["ssim"] for c in comparisons if c["oxi_vs_word"].get("ssim") is not None]
    if vals:
        avg = sum(vals) / len(vals)
        best = max(vals)
        worst = min(vals)
        print(f"  Average: {avg:.4f}")
        print(f"  Best:    {best:.4f}")
        print(f"  Worst:   {worst:.4f}")
        print(f"  Files:   {len(vals)}/{total}")
    else:
        print("  No valid comparisons")
    print("=" * 70)

    # Save report
    report = {
        "total_files": total,
        "dpi": DPI,
        "oxi_results": oxi_results,
        "word_results": word_results,
        "comparisons": comparisons,
        "summary": {
            "avg_ssim": sum(vals) / len(vals) if vals else None,
            "best_ssim": max(vals) if vals else None,
            "worst_ssim": min(vals) if vals else None,
        },
    }
    report_path = OUTPUT_DIR / "pixel_comparison_report.json"
    report_path.write_text(json.dumps(report, indent=2, default=str))
    print(f"\nReport: {report_path}")
    print(f"Side-by-side images: {OUTPUT_DIR / 'comparison'}")


if __name__ == "__main__":
    main()
