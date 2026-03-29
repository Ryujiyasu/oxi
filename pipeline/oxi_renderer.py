"""Oxi CLI (docx-to-pdf) → PyMuPDF → PNG レンダラー"""

import subprocess
import os
import time
import tempfile
from pathlib import Path
from .config import OXI_ROOT, OXI_PNG_DIR, RENDER_DPI

# Max seconds per file before killing oxi
RENDER_TIMEOUT = 30


def render_with_oxi(docx_paths: list[str]) -> dict[str, list[str]]:
    """
    oxi-cli docx-to-pdf で各.docxをPDF化し、PyMuPDFでPNG化する。
    戻り値: {docx_path: [page1.png, page2.png, ...]}
    """
    import fitz

    results = {}
    oxi_bin = os.path.join(OXI_ROOT, "target", "debug", "oxi.exe")
    if not os.path.exists(oxi_bin):
        oxi_bin = os.path.join(OXI_ROOT, "target", "release", "oxi.exe")
    if not os.path.exists(oxi_bin):
        raise FileNotFoundError(
            "oxi binary not found. Run: cargo build --bin oxi"
        )

    for docx_path in docx_paths:
        doc_id  = Path(docx_path).stem
        out_dir = Path(OXI_PNG_DIR) / doc_id
        out_dir.mkdir(parents=True, exist_ok=True)

        # Skip if already rendered
        existing = sorted(out_dir.glob("page_*.png"))
        if existing:
            results[docx_path] = [str(p) for p in existing]
            continue

        # docx → pdf via oxi-cli
        pdf_path = os.path.join(str(out_dir), f"{doc_id}.pdf")

        try:
            proc = subprocess.Popen(
                [oxi_bin, "docx-to-pdf",
                 os.path.abspath(docx_path), pdf_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            try:
                stdout, stderr = proc.communicate(timeout=RENDER_TIMEOUT)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait()
                print(f"[NG] Oxi timeout ({doc_id}): >{RENDER_TIMEOUT}s")
                results[docx_path] = []
                continue

            if proc.returncode != 0:
                err_text = stderr.decode("utf-8", errors="replace")[:300]
                print(f"[NG] Oxi error ({doc_id}):\n{err_text}")
                results[docx_path] = []
                continue

            # pdf → png via PyMuPDF
            doc = fitz.open(pdf_path)
            zoom = RENDER_DPI / 72
            mat = fitz.Matrix(zoom, zoom)
            png_paths = []

            for page_idx in range(len(doc)):
                page = doc[page_idx]
                pix = page.get_pixmap(matrix=mat)
                png_path = str(out_dir / f"page_{page_idx+1:04d}.png")
                pix.save(png_path)
                png_paths.append(png_path)

            doc.close()
            results[docx_path] = png_paths
            print(f"  Oxi: {doc_id} ({len(png_paths)} pages)")

        except Exception as e:
            print(f"[NG] Oxi error ({doc_id}): {e}")
            results[docx_path] = []

        finally:
            if os.path.exists(pdf_path):
                try:
                    os.unlink(pdf_path)
                except PermissionError:
                    time.sleep(1)
                    try:
                        os.unlink(pdf_path)
                    except PermissionError:
                        pass

    print(f"[OK] Oxi rendering done: {len(results)} files")
    return results
