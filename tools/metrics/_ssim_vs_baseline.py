# -*- coding: utf-8 -*-
"""SSIM of the CURRENT build against the committed `ssim_baseline.json`, for a
named subset of bases.

`ssim_ab.py` A/Bs an env flag, which only sees what that flag switches. When a
change RESTRUCTURES a path, the flag's OFF arm still runs the new structure, so
the A/B can report "0 changed" while the rewrite moved pixels. The stored
baseline is the only reference that predates the rewrite, so this compares
against it directly. Word PNGs come from the same cache the pipeline uses.

    python tools/metrics/_ssim_vs_baseline.py albaluna probevert

Uses the DWrite renderer (pipeline.verify's default since S50).
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
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = str(_REPO)
DW = os.environ.get("OXI_DWRITE_EXE") or os.path.join(
    REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
BASELINE = os.path.join(REPO, "pipeline_data", "ssim_baseline.json")
filt = sys.argv[1:]


def ssim2(wpng, opng):
    w = _load_rgb(wpng)
    o = _resize_to_match(_load_rgb(opng), w)
    return ssim(w, o, full=False, channel_axis=2, data_range=255)


def find(base):
    e = Path(DOCS) / (base + ".docx")
    if e.exists():
        return os.path.abspath(str(e))
    c = sorted(p for p in Path(DOCS).glob(base.split("_")[0] + "*.docx")
               if not p.name.startswith("~$"))
    return os.path.abspath(str(c[0])) if c else None


baseline = json.load(open(BASELINE, encoding="utf-8"))
bases = sorted({re.sub(r"_p\d+$", "", n) for n in os.listdir(WORD_PNG_DIR)})
if filt:
    bases = [b for b in bases if any(b.startswith(f) for f in filt)]

rows = []
with tempfile.TemporaryDirectory() as tmp:
    for base in bases:
        d = find(base)
        if not d or base not in baseline:
            continue
        out = Path(tmp) / base
        out.mkdir(parents=True, exist_ok=True)
        subprocess.run([DW, d, str(out / "p"), str(RENDER_DPI)],
                       capture_output=True, timeout=300)
        wdir = Path(WORD_PNG_DIR) / base
        i = 1
        per = []
        while True:
            wp = wdir / f"page_{i:04d}.png"
            op = out / f"p_p{i}.png"
            if not wp.exists():
                break
            if not op.exists():
                per.append((i, None, baseline[base].get(str(i))))
                i += 1
                continue
            key = str(i) if str(i) in baseline[base] else f"{i:04d}"
            per.append((i, ssim2(str(wp), str(op)), baseline[base].get(key)))
            i += 1
        rows.append((base, per))

print(f"{'base':<44} {'pg':>3} {'baseline':>9} {'now':>9} {'delta':>9}")
worst = 0.0
for base, per in rows:
    for pg, now, old in per:
        if now is None:
            print(f"{base:<44} {pg:>3} {old if old is None else f'{old:9.4f}'}    MISSING PAGE")
            continue
        if old is None:
            print(f"{base:<44} {pg:>3} {'-':>9} {now:>9.4f} {'(new)':>9}")
            continue
        d = now - old
        worst = min(worst, d)
        flag = "  <<< REGRESS" if d < -0.001 else ("  >>> improve" if d > 0.001 else "")
        print(f"{base:<44} {pg:>3} {old:>9.4f} {now:>9.4f} {d:>+9.4f}{flag}")
print(f"\nworst per-page delta: {worst:+.4f}")
