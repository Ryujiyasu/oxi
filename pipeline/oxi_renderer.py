"""Oxi Renderer → PNG (DirectWrite/Direct2D default, GDI fallback)"""

import subprocess
import os
from pathlib import Path
from .config import OXI_ROOT, OXI_PNG_DIR, RENDER_DPI

RENDER_TIMEOUT = 300

# DirectWrite is the default per session_50_dwrite_renderer_shipped.md
# (commit 04cc22d): full baseline mean SSIM 0.851286→0.854443 (+0.003157,
# 1.5x sentinel). Word uses DirectWrite internally so this matches Word's
# text engine glyph metrics. Set OXI_USE_GDI=1 to force the legacy GDI
# renderer (e.g., for diagnosing per-doc regressions vs DWrite).
DWRITE_RENDERER = os.path.join(
    OXI_ROOT, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe"
)
GDI_RENDERER = os.path.join(
    OXI_ROOT, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe"
)
USE_GDI = os.environ.get("OXI_USE_GDI", "").lower() in ("1", "true", "yes")
RENDERER_BIN = GDI_RENDERER if USE_GDI else DWRITE_RENDERER
RENDERER_NAME = "GDI" if USE_GDI else "DWrite"


def render_with_oxi(docx_paths: list[str]) -> dict[str, list[str]]:
    """Render each .docx via the active Oxi renderer (DWrite default; GDI
    fallback when OXI_USE_GDI=1).

    Returns: {docx_path: [page1.png, page2.png, ...]}
    """
    results = {}

    if not os.path.exists(RENDERER_BIN):
        crate_dir = "oxi-gdi-renderer" if USE_GDI else "oxi-dwrite-renderer"
        raise FileNotFoundError(
            f"Oxi {RENDERER_NAME} renderer not found: {RENDERER_BIN}\n"
            f"Build with: cd tools/{crate_dir} && cargo build --release"
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

        timed_out = False
        try:
            proc = subprocess.Popen(
                [RENDERER_BIN,
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
                print(f"[WARN] Oxi {RENDERER_NAME} timeout ({doc_id}): >{RENDER_TIMEOUT}s — using partial output")
                timed_out = True

            if not timed_out and proc.returncode != 0:
                err_text = stderr.decode("utf-8", errors="replace")[:300]
                print(f"[NG] Oxi {RENDERER_NAME} error ({doc_id}):\n{err_text}")
                results[docx_path] = []
                continue

            # Rename whatever pages were produced, even on timeout.
            # Some docs render 70+ pages and can't finish in the timeout, but
            # SSIM comparison only uses min(Word, Oxi) pages, so partial output
            # is better than zero output.
            page_idx = 1
            while True:
                src = out_dir / f"oxi_p{page_idx}.png"
                if not src.exists():
                    break
                dst = out_dir / f"page_{page_idx:04d}.png"
                src.rename(dst)
                page_idx += 1

            # When OXI_MAX_PAGES is set, drop extra pages (Oxi renders all
            # internally; we only want p.1..N for the canary).
            max_pages = int(os.environ.get("OXI_MAX_PAGES", "0") or "0")
            if max_pages > 0:
                for extra in sorted(out_dir.glob("page_*.png"))[max_pages:]:
                    extra.unlink()

            png_paths = sorted(out_dir.glob("page_*.png"))
            results[docx_path] = [str(p) for p in png_paths]
            tag = "[WARN] partial" if timed_out else f"  Oxi {RENDERER_NAME}"
            print(f"{tag}: {doc_id} ({len(png_paths)} pages)")

        except Exception as e:
            print(f"[NG] Oxi {RENDERER_NAME} error ({doc_id}): {e}")
            results[docx_path] = []

    print(f"[OK] Oxi {RENDERER_NAME} rendering done: {len(results)} files")
    return results
