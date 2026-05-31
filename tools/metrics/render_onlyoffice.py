"""Batch-render baseline .docx files via OnlyOffice x2t headless to PDF, then
rasterize to PNG at the same DPI as the Oxi/Word/Libra pipelines.

Layout-comparison parity with the Word / Oxi / Libra pipelines:
  - Word side:  pipeline_data/word_png/<doc_id>/page_NNNN.png  (Word COM, 150 DPI)
  - Oxi  side:  pipeline_data/oxi_png/<doc_id>/page_NNNN.png   (oxi-gdi / oxi-dwrite)
  - Libra side: pipeline_data/libra_png/<doc_id>/page_NNNN.png (soffice -> PDF -> pymupdf)
  - OO   side:  pipeline_data/onlyoffice_png/<doc_id>/page_NNNN.png (x2t -> PDF -> pymupdf)

Also caches:
  - pipeline_data/onlyoffice_pdf/<doc_id>.pdf

Open-rate report (crash-free %):
  - pipeline_data/onlyoffice_open_report.json

OnlyOffice x2t setup note (2026-05-31):
  The installed x2t (C:\\Program Files\\ONLYOFFICE\\DesktopEditors\\converter)
  reads DoctRenderer.config from its own exe dir, and that config points at a
  ../editors/sdkjs/common/AllFonts.js that does NOT ship (the real font cache
  lives in %LOCALAPPDATA%\\ONLYOFFICE\\DesktopEditors\\data\\fonts). Program
  Files is not writable without admin, so we run x2t from a writable COPY at
  C:\\tmp\\ooconv whose DoctRenderer.config uses absolute paths to the real
  sdkjs + AllFonts.js. CONV_DIR below points at that copy. Re-create it with:
    robocopy "<install>\\converter" C:\\tmp\\ooconv /E
  then drop the absolute-path DoctRenderer.config (see _write_config()).

Usage (from repo root):
    python tools/metrics/render_onlyoffice.py                # all docs
    python tools/metrics/render_onlyoffice.py --limit 5
    python tools/metrics/render_onlyoffice.py 04b88e         # prefix filter
    python tools/metrics/render_onlyoffice.py --baseline     # only ssim_baseline docs
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DOCS_DIR = REPO_ROOT / "tools" / "golden-test" / "documents" / "docx"
PIPELINE_DATA = REPO_ROOT / "pipeline_data"
OO_PDF_DIR = PIPELINE_DATA / "onlyoffice_pdf"
OO_PNG_DIR = PIPELINE_DATA / "onlyoffice_png"
OPEN_REPORT = PIPELINE_DATA / "onlyoffice_open_report.json"
SSIM_BASELINE = PIPELINE_DATA / "ssim_baseline.json"

# Writable copy of the OnlyOffice converter dir (see module docstring).
CONV_DIR = Path(r"C:\tmp\ooconv")
X2T = CONV_DIR / "x2t.exe"
INSTALL_ROOT = Path(r"C:\Program Files\ONLYOFFICE\DesktopEditors")
ALLFONTS = Path(
    r"C:\Users\ryuji\AppData\Local\ONLYOFFICE\DesktopEditors\data\fonts\AllFonts.js"
)
RENDER_DPI = 150
CONVERT_TIMEOUT_S = 90
RASTERIZE_TIMEOUT_S = 60


def _write_config() -> None:
    """Ensure CONV_DIR/DoctRenderer.config uses absolute paths to the real
    sdkjs + the user font cache (the shipped relative AllFonts.js is absent)."""
    cfg = CONV_DIR / "DoctRenderer.config"
    sdkjs = INSTALL_ROOT / "editors" / "sdkjs"
    content = (
        "<Settings>\n"
        f"<file>{sdkjs / 'common' / 'Native' / 'native.js'}</file>\n"
        f"<file>{sdkjs / 'common' / 'Native' / 'jquery_native.js'}</file>\n"
        f"<allfonts>{ALLFONTS}</allfonts>\n"
        f"<file>{INSTALL_ROOT / 'editors' / 'web-apps' / 'vendor' / 'xregexp' / 'xregexp-all-min.js'}</file>\n"
        f"<sdkjs>{sdkjs}</sdkjs>\n"
        f"<dictionaries>{INSTALL_ROOT / 'dictionaries'}</dictionaries>\n"
        "</Settings>\n"
    )
    cfg.write_text(content, encoding="utf-8")


def discover_docs(prefix: str | None, baseline_only: bool, limit: int) -> list[Path]:
    if not DOCS_DIR.is_dir():
        sys.exit(f"baseline docx dir not found: {DOCS_DIR}")
    paths = sorted(DOCS_DIR.glob("*.docx"))

    if baseline_only and SSIM_BASELINE.is_file():
        with SSIM_BASELINE.open(encoding="utf-8") as f:
            baseline_ids = set(json.load(f).keys())
        paths = [p for p in paths if p.stem in baseline_ids]

    if prefix:
        paths = [p for p in paths if p.stem.startswith(prefix)]

    if limit > 0:
        paths = paths[:limit]
    return paths


def convert_to_pdf(docx_abs: Path, pdf_out: Path) -> tuple[bool, str]:
    """Run x2t (two-arg form) from the writable converter copy. Returns (ok, err)."""
    pdf_out.parent.mkdir(parents=True, exist_ok=True)
    # x2t writes directly to the requested output path; PDF format is inferred
    # from the .pdf extension.
    cmd = [str(X2T), str(docx_abs), str(pdf_out)]
    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True,
            encoding="utf-8", errors="replace",
            cwd=str(CONV_DIR),
            timeout=CONVERT_TIMEOUT_S,
        )
    except subprocess.TimeoutExpired:
        return False, f"x2t timeout after {CONVERT_TIMEOUT_S}s"
    if proc.returncode != 0:
        tail = (proc.stdout or proc.stderr or "")[-200:]
        return False, f"x2t rc={proc.returncode} {tail}"
    if not pdf_out.is_file():
        return False, f"x2t produced no PDF (stdout={proc.stdout[-200:]})"
    return True, ""


def rasterize_pdf(pdf_path: Path, png_dir: Path, dpi: int) -> tuple[bool, str, int]:
    """Rasterize each page to png_dir/page_NNNN.png. Returns (ok, err, n_pages)."""
    png_dir.mkdir(parents=True, exist_ok=True)
    for old in png_dir.glob("page_*.png"):
        try:
            old.unlink()
        except OSError:
            pass
    try:
        import fitz  # pymupdf
    except ImportError:
        return False, "pymupdf not installed", 0
    try:
        with fitz.open(str(pdf_path)) as pdf:
            n_pages = pdf.page_count
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            for i, page in enumerate(pdf, start=1):
                pix = page.get_pixmap(matrix=mat, alpha=False)
                pix.save(str(png_dir / f"page_{i:04d}.png"))
        return True, "", n_pages
    except Exception as e:
        return False, f"pymupdf: {e}", 0


def render_one(docx: Path, force: bool) -> dict:
    pdf_out = OO_PDF_DIR / (docx.stem + ".pdf")
    png_dir = OO_PNG_DIR / docx.stem
    rec: dict = {"doc_id": docx.stem}

    t0 = time.time()
    if pdf_out.is_file() and not force:
        rec["pdf_cached"] = True
        rec["convert_ok"] = True
    else:
        ok, err = convert_to_pdf(docx, pdf_out)
        rec["convert_ok"] = ok
        rec["convert_err"] = err if not ok else None
        if not ok:
            rec["elapsed_s"] = round(time.time() - t0, 2)
            return rec

    existing_pngs = sorted(png_dir.glob("page_*.png")) if png_dir.is_dir() else []
    if existing_pngs and not force:
        rec["png_cached"] = True
        rec["rasterize_ok"] = True
        rec["n_pages"] = len(existing_pngs)
    else:
        ok, err, n = rasterize_pdf(pdf_out, png_dir, RENDER_DPI)
        rec["rasterize_ok"] = ok
        rec["rasterize_err"] = err if not ok else None
        rec["n_pages"] = n

    rec["elapsed_s"] = round(time.time() - t0, 2)
    return rec


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("prefix", nargs="?", default=None, help="doc_id prefix filter")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--baseline", action="store_true",
                    help="only docs that appear in ssim_baseline.json")
    ap.add_argument("--force", action="store_true",
                    help="re-convert and re-rasterize even if cached")
    args = ap.parse_args()

    if not X2T.is_file():
        sys.exit(
            f"x2t not found at {X2T}.\n"
            f"Create the writable converter copy first:\n"
            f'  robocopy "{INSTALL_ROOT / "converter"}" {CONV_DIR} /E'
        )
    if not ALLFONTS.is_file():
        sys.exit(f"OnlyOffice font cache missing: {ALLFONTS}")
    _write_config()

    docs = discover_docs(args.prefix, args.baseline, args.limit)
    if not docs:
        sys.exit("no docs matched")

    print(f"# rendering {len(docs)} doc(s) via OnlyOffice x2t (DPI={RENDER_DPI})")
    OO_PDF_DIR.mkdir(parents=True, exist_ok=True)
    OO_PNG_DIR.mkdir(parents=True, exist_ok=True)

    records = []
    t_total = time.time()
    n_ok_convert = 0
    n_ok_rasterize = 0
    for i, docx in enumerate(docs, start=1):
        print(f"[{i}/{len(docs)}] {docx.stem} ...", end=" ", flush=True)
        rec = render_one(docx, force=args.force)
        records.append(rec)
        if rec.get("convert_ok"):
            n_ok_convert += 1
            tag = "cached" if rec.get("pdf_cached") else "converted"
            extra = ""
            if rec.get("rasterize_ok"):
                n_ok_rasterize += 1
                extra = f" -> {rec.get('n_pages', 0)} pages"
            elif rec.get("rasterize_err"):
                extra = f" RASTERIZE_FAIL: {rec['rasterize_err']}"
            print(f"OK ({tag}) {rec['elapsed_s']}s{extra}")
        else:
            print(f"CONVERT_FAIL: {rec.get('convert_err', '?')} ({rec['elapsed_s']}s)")

    elapsed = round(time.time() - t_total, 1)
    summary = {
        "n_total": len(docs),
        "n_convert_ok": n_ok_convert,
        "n_rasterize_ok": n_ok_rasterize,
        "convert_rate": round(n_ok_convert / len(docs), 4),
        "rasterize_rate": round(n_ok_rasterize / len(docs), 4),
        "render_dpi": RENDER_DPI,
        "elapsed_s": elapsed,
        "x2t_path": str(X2T),
        "records": records,
    }
    OPEN_REPORT.parent.mkdir(parents=True, exist_ok=True)
    with OPEN_REPORT.open("w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(
        f"\n# done. convert {n_ok_convert}/{len(docs)} ({summary['convert_rate']:.1%}), "
        f"rasterize {n_ok_rasterize}/{len(docs)} ({summary['rasterize_rate']:.1%}), "
        f"{elapsed}s total -> {OPEN_REPORT}"
    )


if __name__ == "__main__":
    main()
