#!/usr/bin/env python3
"""
Quick comparison: Oxi vs LibreOffice vs Word (ground truth).
Renders basic_test.docx with all three, computes SSIM against Word output.
"""
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
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
DPI = 150
OUT_DIR = Path(__file__).resolve().parent / "oxi_vs_lo"


def pdf_to_png(pdf_path, out_png, dpi=DPI):
    doc = fitz.open(str(pdf_path))
    if len(doc) == 0:
        return False
    page = doc[0]
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    pix.save(str(out_png))
    doc.close()
    return True


def render_word(docx_path, out_png):
    import win32com.client
    tmp = tempfile.mkdtemp(prefix="word_cmp_")
    try:
        pdf_path = Path(tmp) / (docx_path.stem + ".pdf")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
            doc.ExportAsFixedFormat(str(pdf_path.resolve()), 17, False, 0)
            doc.Close(False)
        finally:
            word.Quit()
        if not pdf_path.exists():
            return False
        return pdf_to_png(pdf_path, out_png)
    except Exception as e:
        print(f"  Word error: {e}")
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def render_libreoffice(docx_path, out_png):
    tmp = tempfile.mkdtemp(prefix="lo_cmp_")
    try:
        profile_dir = Path(tmp) / "profile"
        profile_dir.mkdir()
        profile_url = "file:///" + str(profile_dir).replace("\\", "/")
        out_dir = Path(tmp) / "out"
        out_dir.mkdir()
        subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "pdf",
             "--outdir", str(out_dir),
             f"-env:UserInstallation={profile_url}",
             str(docx_path)],
            capture_output=True, timeout=60,
        )
        pdf_path = out_dir / (docx_path.stem + ".pdf")
        if not pdf_path.exists():
            return False
        return pdf_to_png(pdf_path, out_png)
    except Exception as e:
        print(f"  LO error: {e}")
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def render_oxi(docx_path, out_png):
    oxi_cli = WEB_ROOT / "target" / "release" / "oxi.exe"
    if not oxi_cli.exists():
        oxi_cli = WEB_ROOT / "target" / "release" / "oxi"
    if not oxi_cli.exists():
        print("  oxi-cli not found. Run: cargo build --release -p oxi-cli")
        return False
    tmp = tempfile.mkdtemp(prefix="oxi_cmp_")
    try:
        pdf_path = Path(tmp) / (docx_path.stem + ".pdf")
        result = subprocess.run(
            [str(oxi_cli), "docx-to-pdf", str(docx_path), str(pdf_path)],
            capture_output=True, timeout=60,
        )
        if not pdf_path.exists():
            print(f"  Oxi error: {(result.stderr or b'').decode()[:200]}")
            return False
        return pdf_to_png(pdf_path, out_png)
    except Exception as e:
        print(f"  Oxi error: {e}")
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def compare_images(path1, path2):
    img1 = cv2.imread(str(path1))
    img2 = cv2.imread(str(path2))
    if img1 is None or img2 is None:
        return None, None
    h = min(img1.shape[0], img2.shape[0])
    w = min(img1.shape[1], img2.shape[1])
    img1 = cv2.resize(img1, (w, h))
    img2 = cv2.resize(img2, (w, h))
    gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
    gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
    score, _ = ssim(gray1, gray2, full=True)
    diff = cv2.absdiff(img1, img2)
    diff_px = np.count_nonzero(cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY))
    match_ratio = 1.0 - diff_px / (h * w)
    return round(score, 4), round(match_ratio, 4)


def create_comparison(images, labels, out_path):
    imgs = []
    for label in labels:
        path = images.get(label)
        if path and Path(path).exists():
            img = cv2.imread(str(path))
            if img is not None:
                h, w = img.shape[:2]
                labeled = np.ones((h + 50, w, 3), dtype=np.uint8) * 255
                cv2.putText(labeled, label, (10, 35), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 0), 2)
                labeled[50:50 + h, :w] = img
                imgs.append(labeled)
    if len(imgs) < 2:
        return
    max_h = max(i.shape[0] for i in imgs)
    resized = []
    for img in imgs:
        h, w = img.shape[:2]
        if h < max_h:
            p = np.ones((max_h, w, 3), dtype=np.uint8) * 255
            p[:h, :w] = img
            resized.append(p)
        else:
            resized.append(img)
    combined = np.hstack(resized)
    cv2.imwrite(str(out_path), combined)


def main():
    import random
    random.seed(42)

    # Collect test files: fixtures + sample from golden test docs
    fixtures = WEB_ROOT / "tests" / "fixtures"
    test_files = sorted(fixtures.glob("*.docx"))

    golden_dir = Path(__file__).resolve().parent / "documents" / "docx"
    if golden_dir.exists():
        golden_files = sorted(golden_dir.glob("*.docx"))
        sample = random.sample(golden_files, min(10, len(golden_files)))
        test_files.extend(sample)

    if not test_files:
        print("No .docx files found")
        sys.exit(1)

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    print("=" * 70)
    print("  Oxi vs LibreOffice vs Word (SSIM comparison)")
    print("=" * 70)
    print(f"  Files: {len(test_files)}, DPI: {DPI}")
    print("=" * 70)
    print()

    # Build oxi-cli (skip if binary exists)
    oxi_cli = WEB_ROOT / "target" / "release" / "oxi.exe"
    if not oxi_cli.exists():
        oxi_cli = WEB_ROOT / "target" / "release" / "oxi"
    if not oxi_cli.exists():
        print(">>> Building oxi-cli (release)...", flush=True)
        r = subprocess.run(
            ["cargo", "build", "--release", "-p", "oxi-cli"],
            cwd=str(WEB_ROOT), capture_output=True,
        )
        if r.returncode != 0:
            print(f"  Build failed: {r.stderr.decode()[:300]}")
            sys.exit(1)
        print("    Done.\n", flush=True)
    else:
        print(f">>> oxi-cli found: {oxi_cli}\n", flush=True)

    results = []

    for docx_path in test_files:
        name = docx_path.name
        stem = docx_path.stem
        print(f">>> {name}")

        word_png = OUT_DIR / f"{stem}_word.png"
        oxi_png = OUT_DIR / f"{stem}_oxi.png"
        lo_png = OUT_DIR / f"{stem}_lo.png"

        # Render Word (cached)
        if not word_png.exists():
            print("  Rendering Word...", end="", flush=True)
            ok = render_word(docx_path, word_png)
            print(f" {'OK' if ok else 'FAIL'}")
        else:
            print("  Word: cached")

        # Render Oxi
        print("  Rendering Oxi...", end="", flush=True)
        oxi_ok = render_oxi(docx_path, oxi_png)
        print(f" {'OK' if oxi_ok else 'FAIL'}")

        # Render LibreOffice
        print("  Rendering LO...", end="", flush=True)
        lo_ok = render_libreoffice(docx_path, lo_png)
        print(f" {'OK' if lo_ok else 'FAIL'}")

        # Compare
        oxi_ssim, oxi_match = (None, None)
        lo_ssim, lo_match = (None, None)

        if word_png.exists() and oxi_png.exists():
            oxi_ssim, oxi_match = compare_images(word_png, oxi_png)
        if word_png.exists() and lo_png.exists():
            lo_ssim, lo_match = compare_images(word_png, lo_png)

        winner = "-"
        if oxi_ssim is not None and lo_ssim is not None:
            winner = "Oxi" if oxi_ssim >= lo_ssim else "LO"

        result = {
            "file": name,
            "oxi_ssim": oxi_ssim,
            "oxi_match": oxi_match,
            "lo_ssim": lo_ssim,
            "lo_match": lo_match,
            "winner": winner,
        }
        results.append(result)

        print(f"  Oxi vs Word: SSIM={oxi_ssim}  Match={oxi_match}")
        print(f"  LO  vs Word: SSIM={lo_ssim}  Match={lo_match}")
        print(f"  Winner: {winner}")
        print()

        # Side-by-side comparison image
        imgs = {}
        if word_png.exists(): imgs["Word"] = str(word_png)
        if oxi_png.exists(): imgs["Oxi"] = str(oxi_png)
        if lo_png.exists(): imgs["LibreOffice"] = str(lo_png)
        if len(imgs) >= 2:
            create_comparison(imgs, list(imgs.keys()), OUT_DIR / f"{stem}_compare.png")

    # Summary
    print("=" * 70)
    print(f"  {'File':<40} {'Oxi SSIM':>10} {'LO SSIM':>10} {'Winner':>8}")
    print("-" * 70)
    for r in results:
        oxi_s = f"{r['oxi_ssim']:.4f}" if r['oxi_ssim'] is not None else "N/A"
        lo_s = f"{r['lo_ssim']:.4f}" if r['lo_ssim'] is not None else "N/A"
        print(f"  {r['file']:<40} {oxi_s:>10} {lo_s:>10} {r['winner']:>8}")

    oxi_vals = [r["oxi_ssim"] for r in results if r["oxi_ssim"] is not None]
    lo_vals = [r["lo_ssim"] for r in results if r["lo_ssim"] is not None]
    oxi_wins = sum(1 for r in results if r["winner"] == "Oxi")
    lo_wins = sum(1 for r in results if r["winner"] == "LO")

    print("-" * 70)
    if oxi_vals:
        print(f"  Oxi avg SSIM: {sum(oxi_vals)/len(oxi_vals):.4f}")
    if lo_vals:
        print(f"  LO  avg SSIM: {sum(lo_vals)/len(lo_vals):.4f}")
    print(f"  Wins: Oxi={oxi_wins}, LO={lo_wins}")
    print("=" * 70)
    print(f"\nComparison images: {OUT_DIR}")


if __name__ == "__main__":
    main()
