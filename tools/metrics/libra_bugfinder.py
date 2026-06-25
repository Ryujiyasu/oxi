# -*- coding: utf-8 -*-
"""Fresh Oxi-vs-LibreOffice bug-finder (current binary).

For every doc with BOTH word_png and libra_png, render Oxi (current DWrite),
compute per-page SSIM(Oxi,Word) and SSIM(Libra,Word), rank by delta =
libra - oxi. A large positive delta = LibreOffice matches Word but Oxi
doesn't = a FIXABLE Oxi bug (not a Word quirk). This is the S667/S668/S670
pattern, recomputed fresh (the cached libra_vs_oxi_ssim.json reads a stale
oxi_score from ssim_baseline.json).

Usage: python tools/metrics/libra_bugfinder.py
"""
import os, subprocess, tempfile, sys
from pathlib import Path
sys.path.insert(0, r"c:\Users\ryuji\oxi-main")
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8")
REPO = r"c:\Users\ryuji\oxi-main"
DW = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
LIBRA = os.path.join(REPO, "pipeline_data", "libra_png")

def ssim2(a, b):
    w = _load_rgb(a); o = _resize_to_match(_load_rgb(b), w)
    return ssim(w, o, full=False, channel_axis=2, data_range=255)

def find_docx(base):
    e = Path(DOCS) / (base + ".docx")
    if e.exists(): return str(e)
    c = sorted(p for p in Path(DOCS).glob(base.split("_")[0] + "*.docx") if not p.name.startswith("~$"))
    return str(c[0]) if c else None

bases = sorted(d.name for d in Path(WORD_PNG_DIR).iterdir() if d.is_dir()
               and (Path(LIBRA) / d.name).is_dir())
recs = []
with tempfile.TemporaryDirectory() as tmp:
    for i, base in enumerate(bases):
        docx = find_docx(base)
        if not docx: continue
        od = Path(tmp) / base
        od.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.run([DW, docx, str(od / "p"), str(RENDER_DPI)], capture_output=True, timeout=300)
        except Exception:
            continue
        wdir = Path(WORD_PNG_DIR) / base; ldir = Path(LIBRA) / base
        p = 1
        while True:
            wp = wdir / f"page_{p:04d}.png"; lp = ldir / f"page_{p:04d}.png"; op = od / f"p_p{p}.png"
            if not wp.exists(): break
            if op.exists() and lp.exists():
                try:
                    so = ssim2(str(wp), str(op)); sl = ssim2(str(wp), str(lp))
                    recs.append((base, p, so, sl, sl - so))
                except Exception:
                    pass
            p += 1
        if i % 30 == 0: sys.stderr.write(f"  {i}/{len(bases)}\n")

recs.sort(key=lambda r: -r[4])
print(f"{'doc':44} {'pg':>2} {'oxi':>6} {'libra':>6} {'delta':>7}")
for base, p, so, sl, d in recs:
    if d < 0.05: break
    print(f"{base[:44]:44} {p:>2} {so:.3f} {sl:.3f} {d:+.4f}")
import statistics
print(f"\nscored {len(recs)} pages; mean oxi {statistics.mean(r[2] for r in recs):.4f} libra {statistics.mean(r[3] for r in recs):.4f}")
print(f"libra-better pages (delta>0.01): {sum(1 for r in recs if r[4]>0.01)}")
