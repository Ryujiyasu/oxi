# -*- coding: utf-8 -*-
"""Re-render the blind pptx corpus at HEAD and re-score it, deck by deck.

Refreshes `ssim_pptx/oxi_png/<deck>` so the cache once again represents the
CURRENT renderer, and writes `ssim_pptx/_remeasure.json` after every deck so the
run is resumable and a kill costs at most one deck.

Why it matters: a cached render is only a valid control until the next ship. On
2026-08-27 nine fixes landed in a day, and comparing a fresh render against a
cache from before them credited three commits to the newest one (a false
+0.0215 on d31). Refreshing the cache is what keeps the next gate honest.

★The renderer is NOT parallel-safe (`pptx_render_not_parallel_safe`): decks are
rendered strictly one at a time, and nothing else may render while this runs.

Usage:
    python tools/metrics/pptx_remeasure.py            # all decks, resuming
    python tools/metrics/pptx_remeasure.py --report   # print what is stored
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from skimage.metrics import structural_similarity as ssim

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
SSIMD = ROOT / "ssim_pptx"
OUT = SSIMD / "_remeasure.json"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
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


def report(state: dict) -> None:
    rows = [r for r in state.values() if "mean" in r]
    if not rows:
        print("nothing measured yet")
        return
    rows.sort(key=lambda r: r["mean"])
    print(f"decks {len(rows)}   deck-mean {sum(r['mean'] for r in rows)/len(rows):.6f}")
    mins = sorted(rows, key=lambda r: r["min"])
    print(f"\n{'deck':>5}{'mean':>10}{'MIN':>10}{'slide':>7}{'pages':>7}")
    print("-- by MIN --")
    for r in mins[:12]:
        print(f"{r['doc']:>5}{r['mean']:>10.4f}{r['min']:>10.4f}{r['worst_slide']:>7}{r['pages']:>7}")
    print("-- by mean --")
    for r in rows[:12]:
        print(f"{r['doc']:>5}{r['mean']:>10.4f}{r['min']:>10.4f}{r['worst_slide']:>7}{r['pages']:>7}")


def main() -> None:
    exe_mtime = EXE.stat().st_mtime_ns if EXE.exists() else 0
    state = load()
    if "--report" in sys.argv:
        report(state)
        return
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    tmp = SSIMD / "_remeasure_tmp"
    for item in manifest:
        doc = f"{item['idx']:02d}"
        # `--stale` re-does only the decks whose PNGs predate the binary, so a
        # refresh survives being interrupted: each run picks up where the last
        # one stopped instead of starting the 50-deck sweep again. `--force`
        # still redoes everything.
        if "--stale" in sys.argv:
            pngs = sorted((SSIMD / "oxi_png" / doc).glob("slide_s*.png"))
            fresh = pngs and min(q.stat().st_mtime_ns for q in pngs) >= exe_mtime
            if fresh and doc in state and "mean" in state[doc]:
                continue
        elif doc in state and "mean" in state[doc] and "--force" not in sys.argv:
            continue
        src = ROOT / "pptx" / item["local"]
        pdf_path = SSIMD / "ppt_pdf" / f"{doc}.pdf"
        if not src.exists() or not pdf_path.exists():
            continue
        pdf = pymupdf.open(pdf_path)
        shutil.rmtree(tmp, ignore_errors=True)
        tmp.mkdir(parents=True)
        t0 = time.time()
        r = subprocess.run(
            [str(EXE), str(src), str(tmp / "slide"), str(DPI)],
            capture_output=True, timeout=3600,
        )
        pngs = sorted(tmp.glob("slide_s*.png"))
        if r.returncode != 0 or len(pngs) < len(pdf):
            state[doc] = {"doc": doc, "error": f"rc={r.returncode} {len(pngs)}/{len(pdf)} slides"}
            OUT.write_text(json.dumps(state, indent=1), encoding="utf-8")
            print(f"{doc}: {state[doc]['error']}", flush=True)
            continue
        vals = []
        for i in range(len(pdf)):
            p = tmp / f"slide_s{i+1}.png"
            if not p.exists():
                continue
            pix = pdf[i].get_pixmap(dpi=DPI)
            a = np.asarray(Image.frombytes("RGB", (pix.width, pix.height), pix.samples)).astype(float)
            o = Image.open(p).convert("RGB")
            if o.size != (pix.width, pix.height):
                o = o.resize((pix.width, pix.height), Image.LANCZOS)
            vals.append(ssim(a, np.asarray(o).astype(float), channel_axis=2, data_range=255))
        # promote to the cache so the NEXT gate has a HEAD-current control
        dest = SSIMD / "oxi_png" / doc
        shutil.rmtree(dest, ignore_errors=True)
        shutil.move(str(tmp), str(dest))
        state[doc] = {
            "doc": doc, "pages": len(pdf), "scored": len(vals),
            "mean": sum(vals) / len(vals), "min": min(vals),
            "worst_slide": int(np.argmin(vals)) + 1,
            "secs": round(time.time() - t0, 1),
        }
        OUT.write_text(json.dumps(state, indent=1), encoding="utf-8")
        print(f"{doc}: mean {state[doc]['mean']:.6f}  MIN {state[doc]['min']:.4f} "
              f"(s{state[doc]['worst_slide']})  {state[doc]['secs']:.0f}s", flush=True)
    report(state)


if __name__ == "__main__":
    main()
