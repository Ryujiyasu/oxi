"""Oxi GDI Renderer → PNG レンダラー (Windows GDI TextOutW for pixel-accurate comparison)"""

import subprocess
import os
from pathlib import Path
from .config import OXI_ROOT, OXI_PNG_DIR, RENDER_DPI

RENDER_TIMEOUT = 60

# GDI renderer binary path
GDI_RENDERER = os.path.join(OXI_ROOT, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")


def render_with_oxi(docx_paths: list[str]) -> dict[str, list[str]]:
    """
    oxi-gdi-renderer で各.docxをGDI描画しPNG化する。
    戻り値: {docx_path: [page1.png, page2.png, ...]}
    """
    results = {}

    if not os.path.exists(GDI_RENDERER):
        raise FileNotFoundError(
            f"GDI renderer not found: {GDI_RENDERER}\n"
            "Build with: cd tools/oxi-gdi-renderer && cargo build --release"
        )

    for docx_path in docx_paths:
        doc_id = Path(docx_path).stem
        out_dir = Path(OXI_PNG_DIR) / doc_id
        out_dir.mkdir(parents=True, exist_ok=True)

        # Skip if already rendered
        existing = sorted(out_dir.glob("page_*.png"))
        if existing:
            results[docx_path] = [str(p) for p in existing]
            continue

        # GDI renderer outputs: {prefix}_p1.png, {prefix}_p2.png, ...
        prefix = str(out_dir / "oxi")

        try:
            proc = subprocess.Popen(
                [GDI_RENDERER,
                 os.path.abspath(docx_path),
                 prefix,
                 str(RENDER_DPI)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            try:
                stdout, stderr = proc.communicate(timeout=RENDER_TIMEOUT)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait()
                print(f"[NG] Oxi GDI timeout ({doc_id}): >{RENDER_TIMEOUT}s")
                results[docx_path] = []
                continue

            if proc.returncode != 0:
                err_text = stderr.decode("utf-8", errors="replace")[:300]
                print(f"[NG] Oxi GDI error ({doc_id}):\n{err_text}")
                results[docx_path] = []
                continue

            # Rename output files: oxi_p1.png -> page_0001.png
            page_idx = 1
            while True:
                src = out_dir / f"oxi_p{page_idx}.png"
                if not src.exists():
                    break
                dst = out_dir / f"page_{page_idx:04d}.png"
                src.rename(dst)
                page_idx += 1

            png_paths = sorted(out_dir.glob("page_*.png"))
            results[docx_path] = [str(p) for p in png_paths]
            print(f"  Oxi GDI: {doc_id} ({len(png_paths)} pages)")

        except Exception as e:
            print(f"[NG] Oxi GDI error ({doc_id}): {e}")
            results[docx_path] = []

    print(f"[OK] Oxi GDI rendering done: {len(results)} files")
    return results
