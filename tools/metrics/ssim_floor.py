# -*- coding: utf-8 -*-
"""Per-document SSIM at HEAD, worst first -- the Phase 3 floor.

ssim_ab.py answers "did this change help?" and only scores the documents whose
bytes moved. Neither it nor ssim_direct_mean.py answers "which documents are
worst right now", which is what picks the next target. This renders each
word_png base once with the current DWrite build and lists the per-document mean
(and its worst page), ascending.

    python tools/metrics/ssim_floor.py            # all bases
    python tools/metrics/ssim_floor.py 15076 b837 # prefix filter
    OXI_FLOOR_JSON=path.json python ...           # also dump the full table

Uses the same skimage SSIM + RGB load + resize as pipeline.ssim_calculator, so
the numbers line up with the gate.
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
from pipeline.ssim_calculator import _load_rgb, _resize_to_match  # noqa: E402
from skimage.metrics import structural_similarity as ssim  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
REPO = str(_REPO)
DW = os.environ.get("OXI_DWRITE_EXE") or os.path.join(
    REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
filt = sys.argv[1:]


def find(base):
    e = Path(DOCS) / (base + ".docx")
    if e.exists():
        return os.path.abspath(str(e))
    import glob
    c = sorted(p for p in glob.glob(os.path.join(DOCS, base.split("_")[0] + "*.docx"))
               if not os.path.basename(p).startswith("~$"))
    return os.path.abspath(c[0]) if c else None


def main():
    bases = sorted({re.sub(r"_p\d+$", "", n) for n in os.listdir(WORD_PNG_DIR)})
    if filt:
        bases = [b for b in bases if any(b.startswith(f) or f in b for f in filt)]
    rows = []
    seen = set()
    with tempfile.TemporaryDirectory() as tmp:
        for base in bases:
            d = find(base)
            if not d or d in seen:
                continue
            seen.add(d)
            wdir = Path(WORD_PNG_DIR) / base
            if not (wdir / "page_0001.png").exists():
                continue
            out = Path(tmp) / base
            out.mkdir(parents=True, exist_ok=True)
            try:
                subprocess.run([DW, d, str(out / "p"), str(RENDER_DPI)],
                               capture_output=True, timeout=600)
            except subprocess.TimeoutExpired:
                rows.append({"doc": base, "mean": 0.0, "pages": 0, "note": "timeout"})
                continue
            vals, i = [], 1
            while (wdir / ("page_%04d.png" % i)).exists():
                op = out / ("p_p%d.png" % i)
                if not op.exists():
                    break
                try:
                    w = _load_rgb(str(wdir / ("page_%04d.png" % i)))
                    vals.append(ssim(w, _resize_to_match(_load_rgb(str(op)), w),
                                     channel_axis=2, data_range=255))
                except Exception:
                    pass
                i += 1
            if not vals:
                continue
            n_word = len([1 for k in range(1, 10000)
                          if (wdir / ("page_%04d.png" % k)).exists()]) or len(vals)
            rows.append({"doc": base, "mean": sum(vals) / len(vals),
                         "min": min(vals), "pages": len(vals),
                         "word_pages": n_word})
    rows.sort(key=lambda r: r.get("mean", 0.0))
    print("%-46s %-8s %-8s %s" % ("doc", "mean", "worst_pg", "pages(scored/word)"))
    for r in rows:
        print("%-46s %-8.4f %-8.4f %d/%d"
              % (r["doc"][:46], r.get("mean", 0), r.get("min", 0),
                 r.get("pages", 0), r.get("word_pages", 0)))
    if rows:
        print("\nn=%d  corpus mean=%.4f  floor=%.4f (%s)"
              % (len(rows), sum(r["mean"] for r in rows) / len(rows),
                 rows[0]["mean"], rows[0]["doc"]))
    dest = os.environ.get("OXI_FLOOR_JSON")
    if dest:
        Path(dest).write_text(json.dumps(rows, indent=1), encoding="utf-8")
        print("written:", dest)


if __name__ == "__main__":
    main()
