"""Per-page SSIM baseline tool.

For every docx in pipeline_data/golden_per_page/ (which are 1-Word-page
slices of the original golden corpus):
  1. Render via Oxi GDI renderer → PNG
  2. Render via Word COM (CopyAsPicture + EMF→PNG) → PNG
  3. Compute full-page SSIM
  4. Save results table sorted ascending (worst SSIM first)

Output: pipeline_data/per_page_ssim/_summary.json + per-doc entries.
"""
import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

# DirectWrite renderer is the default (commit 04cc22d). Set OXI_USE_GDI=1
# to use the legacy GDI renderer for comparison.
_USE_GDI = os.environ.get("OXI_USE_GDI", "").lower() in ("1", "true", "yes")
RENDERER = (ROOT / ("tools/oxi-gdi-renderer" if _USE_GDI else "tools/oxi-dwrite-renderer")
            / "target" / "release"
            / ("oxi-gdi-renderer.exe" if _USE_GDI else "oxi-dwrite-renderer.exe"))
SRC_DIR = ROOT / "pipeline_data" / "golden_per_page"
OUT_DIR = ROOT / "pipeline_data" / "per_page_ssim"
OXI_PNG_DIR = OUT_DIR / "oxi_png"
WORD_PNG_DIR = OUT_DIR / "word_png"


def render_oxi(docx_path: Path, out_prefix: str) -> Path:
    """Returns p1.png path (we expect 1 page per slice)."""
    res = subprocess.run(
        [str(RENDERER), str(docx_path), out_prefix],
        check=True, capture_output=True, text=True,
    )
    for line in (res.stdout + "\n" + res.stderr).splitlines():
        line = line.strip()
        if line.startswith("Saved "):
            p = line[len("Saved "):].split(" (")[0]
            if p.endswith("_p1.png"):
                return Path(p)
    return None


def render_oxi_batch(docxs):
    """Render a batch of docx files. Returns dict[stem -> p1_png_path]."""
    out = {}
    OXI_PNG_DIR.mkdir(parents=True, exist_ok=True)
    for i, docx in enumerate(docxs):
        prefix = str(OXI_PNG_DIR / docx.stem)
        try:
            p = render_oxi(docx, prefix)
            if p:
                out[docx.stem] = p
        except Exception as e:
            print(f"[oxi {i+1}/{len(docxs)}] {docx.name} FAIL: {e}", file=sys.stderr)
        if (i + 1) % 25 == 0:
            print(f"  oxi: {i+1}/{len(docxs)}")
    return out


def render_word_batch(docxs):
    """Render via pipeline.word_renderer in chunks (restart Word periodically)."""
    from pipeline.word_renderer import render_with_word
    from pipeline import config as pc
    pc_orig = pc.WORD_PNG_DIR
    pc.WORD_PNG_DIR = str(WORD_PNG_DIR)
    try:
        import pipeline.word_renderer as wr
        wr.WORD_PNG_DIR = str(WORD_PNG_DIR)
        result = render_with_word([str(d) for d in docxs])
    finally:
        pc.WORD_PNG_DIR = pc_orig
    out = {}
    for d in docxs:
        pngs = result.get(str(d)) or []
        if pngs:
            out[d.stem] = Path(pngs[0])
    return out


def compute_ssim(word_png: Path, oxi_png: Path) -> float:
    w = Image.open(word_png).convert("L")
    o = Image.open(oxi_png).convert("L")
    if w.size != o.size:
        o = o.resize(w.size)
    return float(ssim(np.array(w), np.array(o), data_range=255))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--skip-oxi", action="store_true", help="reuse existing oxi PNGs")
    ap.add_argument("--skip-word", action="store_true", help="reuse existing word PNGs")
    args = ap.parse_args()

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    docxs = sorted(SRC_DIR.glob("*.docx"))
    if args.limit > 0:
        docxs = docxs[: args.limit]
    print(f"docs: {len(docxs)}")

    # Render Oxi
    if args.skip_oxi:
        print("skipping oxi render, reusing existing")
        oxi_pngs = {}
        for d in docxs:
            p = OXI_PNG_DIR / f"{d.stem}_p1.png"
            if p.exists(): oxi_pngs[d.stem] = p
    else:
        print("rendering oxi...")
        oxi_pngs = render_oxi_batch(docxs)
        print(f"  oxi done: {len(oxi_pngs)} pngs")

    # Render Word
    if args.skip_word:
        print("skipping word render, reusing existing")
        word_pngs = {}
        for d in docxs:
            for cand in [WORD_PNG_DIR / f"{d.stem}_p1.png",
                         WORD_PNG_DIR / d.stem / "page_0001.png"]:
                if cand.exists():
                    word_pngs[d.stem] = cand
                    break
    else:
        print("rendering word (slow)...")
        word_pngs = render_word_batch(docxs)
        print(f"  word done: {len(word_pngs)} pngs")

    # SSIM
    print("computing SSIM...")
    rows = []
    for d in docxs:
        wp = word_pngs.get(d.stem); op = oxi_pngs.get(d.stem)
        if not wp or not op:
            rows.append({"doc": d.stem, "ssim": None, "note": "missing png"})
            continue
        try:
            s = compute_ssim(wp, op)
            rows.append({"doc": d.stem, "ssim": s, "word_png": str(wp), "oxi_png": str(op)})
        except Exception as e:
            rows.append({"doc": d.stem, "ssim": None, "note": str(e)})

    rows.sort(key=lambda r: r.get("ssim") if r.get("ssim") is not None else 1.0)
    summary = {
        "n_docs": len(docxs),
        "n_with_ssim": sum(1 for r in rows if r.get("ssim") is not None),
        "mean_ssim": float(np.mean([r["ssim"] for r in rows if r.get("ssim") is not None])),
        "rows": rows,
    }
    out_path = OUT_DIR / "_summary.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"\nsummary -> {out_path}")
    print(f"  n_docs={summary['n_docs']} with_ssim={summary['n_with_ssim']} mean_ssim={summary['mean_ssim']:.4f}")
    print("\nworst 15:")
    for r in rows[:15]:
        s = r.get("ssim"); s_str = f"{s:.4f}" if s is not None else "  --  "
        print(f"  {s_str}  {r['doc']}")


if __name__ == "__main__":
    main()
