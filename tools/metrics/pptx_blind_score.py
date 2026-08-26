# -*- coding: utf-8 -*-
"""Score the blind pptx set one deck at a time, saving as it goes.

The bundled harness (`pipeline_data/pptx_benchmark/_measure_ssim_pptx.py`)
scores every deck and writes its result once, at the end. A run that is
interrupted therefore loses everything -- and a full pass takes long enough
that three of them in a row were killed before finishing, each time discarding
~40 minutes of SSIM.

This does the same arithmetic per deck and writes after each one, so a run can
be stopped and resumed, or split across several invocations, without losing
work. It renders nothing: point it at a cache the renderer has already filled
(the harness's own `oxi_png/<doc>/slide_sN.png` layout).

★It will REFUSE a deck whose cache is short of the slide count PowerPoint
produced, rather than quietly scoring the slides that happen to exist -- the
partial-cache trap that made a 19-slide deck read as a 4-slide one.

Usage:
    python tools/metrics/pptx_blind_score.py            # every deck, resuming
    python tools/metrics/pptx_blind_score.py 1 12       # only docs 1..12
    python tools/metrics/pptx_blind_score.py --report   # print what is stored
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from skimage.metrics import structural_similarity as ssim

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
SSIM = ROOT / "ssim_pptx"
OUT = SSIM / "_incremental.json"
DPI = 150

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def load() -> dict:
    if OUT.exists():
        try:
            return json.loads(OUT.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save(state: dict) -> None:
    OUT.write_text(json.dumps(state, indent=1, ensure_ascii=False), encoding="utf-8")


def score_deck(doc: str) -> dict | None:
    pdf_path = SSIM / "ppt_pdf" / f"{doc}.pdf"
    png_dir = SSIM / "oxi_png" / doc
    if not pdf_path.exists() or not png_dir.is_dir():
        return None
    pdf = pymupdf.open(pdf_path)
    have = list(png_dir.glob("slide_s*.png"))
    if len(have) < len(pdf):
        return {"doc": doc, "error": f"cache has {len(have)} of {len(pdf)} slides"}
    vals = []
    for i in range(len(pdf)):
        p = png_dir / f"slide_s{i+1}.png"
        if not p.exists():
            continue
        pix = pdf[i].get_pixmap(dpi=DPI)
        a = np.asarray(Image.frombytes("RGB", (pix.width, pix.height), pix.samples)).astype(float)
        o = Image.open(p).convert("RGB")
        if o.size != (pix.width, pix.height):
            o = o.resize((pix.width, pix.height), Image.LANCZOS)
        vals.append(ssim(a, np.asarray(o).astype(float), channel_axis=2, data_range=255))
    if not vals:
        return {"doc": doc, "error": "no comparable slides"}
    return {
        "doc": doc,
        "pages": len(pdf),
        "scored": len(vals),
        "mean": sum(vals) / len(vals),
        "min": min(vals),
        "worst_slide": int(np.argmin(vals)) + 1,
    }


def report(state: dict) -> None:
    rows = [r for r in state.values() if "mean" in r]
    if not rows:
        print("nothing scored yet")
        return
    rows.sort(key=lambda r: r["mean"])
    total = sum(r["mean"] for r in rows) / len(rows)
    print(f"scored {len(rows)} decks   deck-mean {total:.6f}\n")
    print(f"{'doc':>4}{'mean':>10}{'worst':>10}{'slide':>7}{'pages':>7}")
    for r in rows[:14]:
        print(f"{r['doc']:>4}{r['mean']:>10.4f}{r['min']:>10.4f}{r['worst_slide']:>7}{r['pages']:>7}")
    bad = [r for r in state.values() if "error" in r]
    for r in bad:
        print(f"  {r['doc']}: {r['error']}")


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    state = load()
    if "--report" in sys.argv:
        report(state)
        return
    lo, hi = (int(args[0]), int(args[1])) if len(args) >= 2 else (1, 50)
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    for item in manifest:
        idx = item["idx"]
        if not (lo <= idx <= hi):
            continue
        doc = f"{idx:02d}"
        if doc in state and "mean" in state[doc] and "--force" not in sys.argv:
            continue
        r = score_deck(doc)
        if r is None:
            continue
        state[doc] = r
        save(state)
        if "mean" in r:
            print(f"{doc}: {r['mean']:.6f}  worst s{r['worst_slide']}={r['min']:.4f}", flush=True)
        else:
            print(f"{doc}: {r['error']}", flush=True)
    report(state)


if __name__ == "__main__":
    main()
