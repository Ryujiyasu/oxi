#!/usr/bin/env python3
"""
Pixel-level comparison: Oxi vs LibreOffice vs OnlyOffice
Renders each .docx file with all three tools and compares output images using SSIM.

Requirements: playwright, PyMuPDF, scikit-image, opencv-python, numpy, pillow
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
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
ONLYOFFICE = r"C:\Program Files\ONLYOFFICE\DesktopEditors\DesktopEditors.exe"
X2T = r"C:\Program Files\ONLYOFFICE\DesktopEditors\converter\x2t.exe"
WEB_ROOT = Path(__file__).resolve().parent.parent.parent  # oxi-1/
WEB_URL = "http://localhost:8080/web/"
PRINT_PREVIEW_URL = "http://localhost:8080/web/print-preview.html"
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
        print(f"    PDF→PNG error: {e}")
        return False


def render_libreoffice(docx_path: Path, out_png: Path) -> dict:
    """Render docx via LibreOffice headless → PDF → PNG."""
    start = time.time()
    tmp = tempfile.mkdtemp(prefix="lo_pixel_")
    profile_dir = os.path.join(tmp, "profile")
    pdf_dir = os.path.join(tmp, "pdf")
    os.makedirs(profile_dir)
    os.makedirs(pdf_dir)
    profile_url = "file:///" + profile_dir.replace("\\", "/")

    try:
        result = subprocess.run(
            [SOFFICE, "--headless", "--norestore", "--nologo",
             "--convert-to", "pdf",
             "--outdir", pdf_dir,
             f"-env:UserInstallation={profile_url}",
             str(docx_path)],
            capture_output=True, text=True, timeout=TIMEOUT_SEC,
        )
        pdf_path = Path(pdf_dir) / (docx_path.stem + ".pdf")
        if not pdf_path.exists():
            return {"success": False, "error": "No PDF output", "time_ms": int((time.time() - start) * 1000)}

        if not pdf_to_png(pdf_path, out_png):
            return {"success": False, "error": "PDF→PNG failed", "time_ms": int((time.time() - start) * 1000)}

        return {"success": True, "error": None, "time_ms": int((time.time() - start) * 1000)}
    except subprocess.TimeoutExpired:
        subprocess.run(["taskkill", "/f", "/im", "soffice.exe"], capture_output=True, timeout=5)
        subprocess.run(["taskkill", "/f", "/im", "soffice.bin"], capture_output=True, timeout=5)
        return {"success": False, "error": "Timeout", "time_ms": int((time.time() - start) * 1000)}
    except Exception as e:
        return {"success": False, "error": str(e)[:200], "time_ms": int((time.time() - start) * 1000)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def render_onlyoffice(docx_path: Path, out_png: Path) -> dict:
    """Render docx via OnlyOffice x2t → PDF → PNG."""
    start = time.time()
    tmp = tempfile.mkdtemp(prefix="oo_pixel_")

    try:
        pdf_path = Path(tmp) / (docx_path.stem + ".pdf")

        # Use x2t converter (direct file-to-file conversion)
        result = subprocess.run(
            [X2T, str(docx_path), str(pdf_path)],
            capture_output=True, text=True, timeout=TIMEOUT_SEC,
        )
        elapsed = int((time.time() - start) * 1000)

        if not pdf_path.exists():
            stderr = (result.stderr or result.stdout or "").strip()[:200]
            return {"success": False, "error": stderr or "No PDF output", "time_ms": elapsed}

        if not pdf_to_png(pdf_path, out_png):
            return {"success": False, "error": "PDF→PNG failed", "time_ms": elapsed}

        return {"success": True, "error": None, "time_ms": elapsed}
    except subprocess.TimeoutExpired:
        return {"success": False, "error": "Timeout", "time_ms": int((time.time() - start) * 1000)}
    except Exception as e:
        return {"success": False, "error": str(e)[:200], "time_ms": int((time.time() - start) * 1000)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def render_oxi_batch(docx_files: list, output_dir: Path) -> dict:
    """Render all docx files via Oxi web UI using Playwright.
    Returns dict of {filename: {success, error, time_ms, png_path}}.
    """
    from playwright.sync_api import sync_playwright

    results = {}
    server_process = None

    try:
        # Check if server is already running
        import urllib.request
        server_running = False
        try:
            urllib.request.urlopen(WEB_URL, timeout=3)
            server_running = True
            print("    Using existing server on port 8080", flush=True)
        except Exception:
            pass

        if not server_running:
            server_process = subprocess.Popen(
                [sys.executable, "-m", "http.server", "8080"],
                cwd=str(WEB_ROOT),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            time.sleep(2)

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1280, "height": 2000})

            for docx_path in docx_files:
                start = time.time()
                fname = docx_path.name
                out_png = output_dir / "oxi" / f"{docx_path.stem}.png"
                out_png.parent.mkdir(parents=True, exist_ok=True)

                try:
                    page = context.new_page()
                    page.goto(PRINT_PREVIEW_URL, wait_until="networkidle", timeout=30000)
                    # Wait for WASM to be ready
                    page.wait_for_function("() => window.__oxiWasmReady === true", timeout=15000)

                    # Load file via fileInput
                    page.locator("#fileInput").set_input_files(str(docx_path))
                    # Wait for render completion
                    page.wait_for_function("() => window.__oxiRenderDone === true", timeout=30000)
                    page.wait_for_timeout(800)  # extra time for async images

                    # Get timing info from the page
                    layout_ms = page.evaluate("() => window.__oxiLayoutMs || 0")
                    total_ms = page.evaluate("() => window.__oxiTotalMs || 0")

                    # Screenshot first page canvas
                    canvas = page.locator("#page-0")
                    if canvas.count() > 0:
                        canvas.screenshot(path=str(out_png))
                        elapsed = int((time.time() - start) * 1000)
                        results[fname] = {
                            "success": True, "error": None,
                            "time_ms": elapsed,
                            "layout_ms": layout_ms,
                            "render_ms": total_ms,
                            "png": str(out_png),
                        }
                    else:
                        elapsed = int((time.time() - start) * 1000)
                        results[fname] = {"success": False, "error": "No page canvas", "time_ms": elapsed, "png": None}
                except Exception as e:
                    elapsed = int((time.time() - start) * 1000)
                    results[fname] = {"success": False, "error": str(e)[:200], "time_ms": elapsed, "png": None}
                finally:
                    page.close()

            browser.close()
    finally:
        if server_process:
            server_process.terminate()
            try:
                server_process.wait(timeout=5)
            except Exception:
                pass

    return results


def compare_images(img1_path: Path, img2_path: Path) -> dict:
    """Compare two images using SSIM and pixel diff metrics."""
    try:
        img1 = cv2.imread(str(img1_path))
        img2 = cv2.imread(str(img2_path))

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
            img = cv2.imread(str(path))
            if img is not None:
                # Add label
                h, w = img.shape[:2]
                labeled = np.zeros((h + 40, w, 3), dtype=np.uint8)
                labeled[:] = (255, 255, 255)
                cv2.putText(labeled, label, (10, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 0), 2)
                labeled[40:40 + h, :w] = img
                imgs.append(labeled)
            else:
                imgs.append(np.zeros((400, 300, 3), dtype=np.uint8))
        else:
            # Placeholder
            placeholder = np.ones((400, 300, 3), dtype=np.uint8) * 200
            cv2.putText(placeholder, f"{label}: N/A", (10, 200), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 0), 1)
            imgs.append(placeholder)

    if not imgs:
        return

    # Resize all to same height
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
    cv2.imwrite(str(out_path), combined)


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
    print("=" * 70)
    print("  Pixel-Level Comparison: Oxi vs LibreOffice vs OnlyOffice")
    print("=" * 70)
    print(f"  Files: {total}")
    print(f"  DPI: {DPI}")
    print(f"  LibreOffice: {'FOUND' if Path(SOFFICE).exists() else 'NOT FOUND'}")
    print(f"  OnlyOffice:  {'FOUND' if Path(X2T).exists() else 'NOT FOUND'} (x2t converter)")
    print("=" * 70)
    print()

    has_lo = Path(SOFFICE).exists()
    has_oo = Path(X2T).exists()

    # Prepare output directories
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for sub in ["oxi", "libreoffice", "onlyoffice", "comparison"]:
        (OUTPUT_DIR / sub).mkdir(parents=True, exist_ok=True)

    # Step 1: Render all with Oxi (batch via Playwright)
    print(">>> Rendering with Oxi (Playwright)...", flush=True)
    oxi_results = render_oxi_batch(test_files, OUTPUT_DIR)
    oxi_ok = sum(1 for r in oxi_results.values() if r["success"])
    print(f"    Oxi: {oxi_ok}/{total} rendered\n", flush=True)

    # Step 2: Render with LibreOffice and OnlyOffice
    lo_results = {}
    oo_results = {}

    for i, docx_path in enumerate(test_files):
        fname = docx_path.name
        stem = docx_path.stem
        print(f"[{i + 1:>3}/{total:>3}] {fname[:50]}", end="", flush=True)

        # LibreOffice
        if has_lo:
            lo_png = OUTPUT_DIR / "libreoffice" / f"{stem}.png"
            lo_res = render_libreoffice(docx_path, lo_png)
            lo_results[fname] = {**lo_res, "png": str(lo_png) if lo_res["success"] else None}
            lo_status = "OK" if lo_res["success"] else "FAIL"
            print(f"  LO:{lo_status}", end="", flush=True)
        time.sleep(0.5)

        # OnlyOffice
        if has_oo:
            oo_png = OUTPUT_DIR / "onlyoffice" / f"{stem}.png"
            oo_res = render_onlyoffice(docx_path, oo_png)
            oo_results[fname] = {**oo_res, "png": str(oo_png) if oo_res["success"] else None}
            oo_status = "OK" if oo_res["success"] else "FAIL"
            print(f"  OO:{oo_status}", end="", flush=True)

        print(flush=True)

    # Step 3: Compare images
    print("\n>>> Computing SSIM comparisons...\n", flush=True)
    comparisons = []

    for docx_path in test_files:
        fname = docx_path.name
        stem = docx_path.stem

        oxi_png = OUTPUT_DIR / "oxi" / f"{stem}.png"
        lo_png = OUTPUT_DIR / "libreoffice" / f"{stem}.png"
        oo_png = OUTPUT_DIR / "onlyoffice" / f"{stem}.png"

        comp = {"filename": fname}

        # Oxi vs LibreOffice
        if oxi_png.exists() and lo_png.exists():
            comp["oxi_vs_lo"] = compare_images(oxi_png, lo_png)
        else:
            comp["oxi_vs_lo"] = {"ssim": None, "error": "Missing image(s)"}

        # Oxi vs OnlyOffice
        if oxi_png.exists() and oo_png.exists():
            comp["oxi_vs_oo"] = compare_images(oxi_png, oo_png)
        else:
            comp["oxi_vs_oo"] = {"ssim": None, "error": "Missing image(s)"}

        # LibreOffice vs OnlyOffice
        if lo_png.exists() and oo_png.exists():
            comp["lo_vs_oo"] = compare_images(lo_png, oo_png)
        else:
            comp["lo_vs_oo"] = {"ssim": None, "error": "Missing image(s)"}

        comparisons.append(comp)

        # Create side-by-side
        side_images = {}
        if oxi_png.exists():
            side_images["Oxi"] = str(oxi_png)
        if lo_png.exists():
            side_images["LibreOffice"] = str(lo_png)
        if oo_png.exists():
            side_images["OnlyOffice"] = str(oo_png)
        if len(side_images) >= 2:
            create_side_by_side(side_images,
                                OUTPUT_DIR / "comparison" / f"{stem}_compare.png",
                                list(side_images.keys()))

        # Print SSIM
        oxi_lo_ssim = comp["oxi_vs_lo"].get("ssim", "-")
        oxi_oo_ssim = comp["oxi_vs_oo"].get("ssim", "-")
        lo_oo_ssim = comp["lo_vs_oo"].get("ssim", "-")
        print(f"  {fname[:45]:45}  Oxi-LO:{oxi_lo_ssim}  Oxi-OO:{oxi_oo_ssim}  LO-OO:{lo_oo_ssim}", flush=True)

    # Summary
    print("\n" + "=" * 70)
    print("  SSIM SUMMARY (1.0 = identical)")
    print("=" * 70)

    def avg_ssim(key):
        vals = [c[key]["ssim"] for c in comparisons if c[key].get("ssim") is not None]
        return sum(vals) / len(vals) if vals else None

    pairs = [("oxi_vs_lo", "Oxi vs LibreOffice"), ("oxi_vs_oo", "Oxi vs OnlyOffice"), ("lo_vs_oo", "LO vs OnlyOffice")]
    for key, label in pairs:
        avg = avg_ssim(key)
        if avg is not None:
            print(f"  {label:25} avg SSIM: {avg:.4f}")
        else:
            print(f"  {label:25} avg SSIM: N/A")
    print("=" * 70)

    # Save report
    report = {
        "total_files": total,
        "dpi": DPI,
        "oxi_results": oxi_results,
        "libreoffice_results": lo_results,
        "onlyoffice_results": oo_results,
        "comparisons": comparisons,
        "summary": {
            key: avg_ssim(key) for key, _ in pairs
        },
    }
    report_path = OUTPUT_DIR / "pixel_comparison_report.json"
    report_path.write_text(json.dumps(report, indent=2, default=str))
    print(f"\nReport: {report_path}")
    print(f"Side-by-side images: {OUTPUT_DIR / 'comparison'}")


if __name__ == "__main__":
    main()
