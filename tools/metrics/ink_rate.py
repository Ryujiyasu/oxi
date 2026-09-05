# -*- coding: utf-8 -*-
"""Ink coverage next to SSIM: how much of each page is ink, in Word and in Oxi,
and how much of that ink overlaps.

SSIM is a windowed structural correlation: a paragraph drawn twice or dropped
outright moves it a little, an anti-aliasing difference moves it a lot. The
two numbers here are blind to anti-aliasing and see exactly the former:

  ink_w / ink_o   fraction of the page's pixels that are ink (luma < INK_THR)
  delta           ink_o - ink_w, as a fraction of ink_w (+ = Oxi draws more)
  iou             |ink_w & ink_o| / |ink_w | ink_o|  (positional agreement)

Usage: python tools/metrics/ink_rate.py [base-prefix ...] [--pages] [--thr=128] [--dilate=2]
The IoU is taken on masks dilated by --dilate pixels (default 2 at 150dpi ~ 1pt), so a
sub-pixel/1px placement difference does not read as a miss; --dilate=0 is the raw mask.
Renders each word_png base with the DWrite renderer (default env), compares
page by page against pipeline_data/word_png/<base>/page_NNNN.png, prints one
line per document (worst page named) and writes
pipeline_data/ink_rate/_summary.json.
"""
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO))
from pipeline.config import WORD_PNG_DIR, RENDER_DPI  # noqa: E402
import numpy as np  # noqa: E402
from PIL import Image  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
DW = os.environ.get("OXI_DWRITE_EXE") or str(_REPO / "tools" / "oxi-dwrite-renderer" / "target" / "release" / "oxi-dwrite-renderer.exe")
DOCS = _REPO / "tools" / "golden-test" / "documents" / "docx"
OUT = _REPO / "pipeline_data" / "ink_rate"

args = [a for a in sys.argv[1:] if not a.startswith("--")]
show_pages = "--pages" in sys.argv
INK_THR = next((int(a.split("=")[1]) for a in sys.argv[1:] if a.startswith("--thr=")), 128)
DILATE = next((int(a.split("=")[1]) for a in sys.argv[1:] if a.startswith("--dilate=")), 2)


def dilate(mask, n):
    if n <= 0:
        return mask
    out = mask.copy()
    for dy in range(-n, n + 1):
        for dx in range(-n, n + 1):
            if dy == 0 and dx == 0:
                continue
            sh = np.roll(np.roll(mask, dy, axis=0), dx, axis=1)
            out |= sh
    return out


def find(base):
    e = DOCS / (base + ".docx")
    if e.exists():
        return str(e)
    c = sorted(p for p in DOCS.glob(base.split("_")[0] + "*.docx") if not p.name.startswith("~$"))
    return str(c[0]) if c else None


def render(docx, outdir):
    Path(outdir).mkdir(parents=True, exist_ok=True)
    subprocess.run([DW, docx, str(Path(outdir) / "p"), str(RENDER_DPI)], capture_output=True, timeout=300)
    ps = []
    i = 1
    while (Path(outdir) / f"p_p{i}.png").exists():
        ps.append(str(Path(outdir) / f"p_p{i}.png"))
        i += 1
    return ps


def ink_mask(path, size=None):
    im = Image.open(path).convert("L")
    if size is not None and im.size != size:
        im = im.resize(size, Image.BILINEAR)
    return np.asarray(im) < INK_THR


def compare(wpng, opng):
    w = ink_mask(wpng)
    o = ink_mask(opng, size=(w.shape[1], w.shape[0]))
    ink_w = float(w.mean())
    ink_o = float(o.mean())
    wd = dilate(w, DILATE)
    od = dilate(o, DILATE)
    union = np.logical_or(wd, od).sum()
    inter = np.logical_and(wd, od).sum()
    iou = float(inter / union) if union else 1.0
    delta = (ink_o - ink_w) / ink_w if ink_w > 0 else 0.0
    return ink_w, ink_o, delta, iou


def main():
    bases = sorted({re.sub(r"_p\d+$", "", n) for n in os.listdir(WORD_PNG_DIR)})
    if args:
        bases = [b for b in bases if any(b.startswith(f) for f in args)]
    OUT.mkdir(parents=True, exist_ok=True)
    rows = []
    seen = set()
    with tempfile.TemporaryDirectory() as tmp:
        for base in bases:
            d = find(base)
            if not d or d in seen:
                continue
            seen.add(d)
            pages = render(d, Path(tmp) / Path(d).stem)
            wdir = Path(WORD_PNG_DIR) / base
            per = []
            i = 1
            while True:
                wp = wdir / f"page_{i:04d}.png"
                if not wp.exists() or i > len(pages):
                    break
                try:
                    per.append((i,) + compare(str(wp), pages[i - 1]))
                except Exception as e:  # noqa: BLE001
                    print("  %s p%d: %s" % (base, i, str(e)[:60]))
                i += 1
            if not per:
                continue
            mean_w = sum(p[1] for p in per) / len(per)
            mean_o = sum(p[2] for p in per) / len(per)
            mean_iou = sum(p[4] for p in per) / len(per)
            worst = min(per, key=lambda p: p[4])
            rows.append({"base": base, "pages": len(per), "word_pages": i - 1, "oxi_pages": len(pages),
                         "ink_w": mean_w, "ink_o": mean_o, "delta": (mean_o - mean_w) / mean_w if mean_w else 0.0,
                         "iou": mean_iou, "worst_page": worst[0], "worst_iou": worst[4], "worst_delta": worst[3],
                         "per_page": [{"page": p[0], "ink_w": p[1], "ink_o": p[2], "delta": p[3], "iou": p[4]} for p in per]})
            if show_pages:
                for p in per:
                    print("  %-28s p%-3d ink_w=%.4f ink_o=%.4f delta=%+6.1f%% iou=%.3f" % (base, p[0], p[1], p[2], 100 * p[3], p[4]))
    rows.sort(key=lambda r: r["iou"])
    print("%-30s %5s  %7s %7s %8s  %6s  worst" % ("base", "pages", "ink_w", "ink_o", "delta", "iou"))
    for r in rows:
        print("%-30s %2d/%-2d  %.4f %.4f %+7.1f%%  %.3f  p%d iou=%.3f delta=%+.1f%%" % (
            r["base"][:30], r["pages"], r["oxi_pages"], r["ink_w"], r["ink_o"], 100 * r["delta"], r["iou"],
            r["worst_page"], r["worst_iou"], 100 * r["worst_delta"]))
    if rows:
        print("mean iou=%.4f  mean |delta|=%.2f%%  docs=%d" % (sum(r["iou"] for r in rows) / len(rows), 100 * sum(abs(r["delta"]) for r in rows) / len(rows), len(rows)))
    (OUT / "_summary.json").write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
