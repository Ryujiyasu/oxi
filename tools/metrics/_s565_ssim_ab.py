# -*- coding: utf-8 -*-
"""S562B — SSIM A/B sentinel (FIXED: bypass render_with_oxi's skip-if-cached).
Render each target doc with the DWrite binary directly to per-mode temp dirs —
ON (default) and OFF (OXI_S565_DISABLE=1) — then SSIM each page vs the cached
word_png both ways and report ON−OFF. Baseline-independent (ssim_baseline.json
is stale per S558).

Usage: python _s559_ssim_ab.py <docids.txt | doc_id...>
"""
import os
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from pipeline.config import WORD_PNG_DIR, RENDER_DPI  # noqa: E402
from pipeline.ssim_calculator import calculate_ssim  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
REPO = r"c:\Users\ryuji\oxi-main"
DW = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
DOCS_DIR = os.path.join(REPO, "tools", "golden-test", "documents", "docx")

args = sys.argv[1:]
targets = []
for a in args:
    if os.path.isfile(a):
        targets += [l.strip() for l in open(a) if l.strip()]
    else:
        targets.append(a)
# drop the synthetic collection ids (no per-doc word_png)
targets = [t for t in targets if t not in ("gen", "gen2", "pixel", "test")]

docx_paths = []
for t in targets:
    cands = sorted(p for p in Path(DOCS_DIR).glob(t + "*.docx") if not p.name.startswith("~$"))
    if cands:
        docx_paths.append(os.path.abspath(str(cands[0])))
    else:
        print("missing docx:", t)


def render(disable, outroot):
    env = dict(os.environ)
    if disable:
        env["OXI_S565_DISABLE"] = "1"
    else:
        env.pop("OXI_S565_DISABLE", None)
    oxi = {}
    for dp in docx_paths:
        d = Path(outroot) / Path(dp).stem
        d.mkdir(parents=True, exist_ok=True)
        subprocess.run([DW, os.path.abspath(dp), str(d / "p"), str(RENDER_DPI)],
                       capture_output=True, timeout=300, env=env)
        # DWrite outputs p_p1.png, p_p2.png, ...
        pages = []
        i = 1
        while (d / f"p_p{i}.png").exists():
            pages.append(str(d / f"p_p{i}.png"))
            i += 1
        oxi[dp] = pages
    return oxi


def word_map():
    w = {}
    for dp in docx_paths:
        wd = Path(WORD_PNG_DIR) / Path(dp).stem
        if wd.exists():
            w[dp] = sorted(str(p) for p in wd.glob("page_*.png"))
    return w


word = word_map()
with tempfile.TemporaryDirectory(prefix="s559_off_") as off_dir, \
     tempfile.TemporaryDirectory(prefix="s559_on_") as on_dir:
    print("rendering OFF (OXI_S565_DISABLE=1)...")
    off_oxi = render(True, off_dir)
    print("rendering ON (default)...")
    on_oxi = render(False, on_dir)
    print("scoring OFF...")
    off_s = {(s["doc_id"], int(s["page"])): s["ssim_score"]
             for s in calculate_ssim(word, off_oxi, skip_heatmap=True)}
    print("scoring ON...")
    on_s = {(s["doc_id"], int(s["page"])): s["ssim_score"]
            for s in calculate_ssim(word, on_oxi, skip_heatmap=True)}

imp = reg = 0
gain = loss = 0.0
per_doc = {}
for k in sorted(set(off_s) & set(on_s)):
    d = on_s[k] - off_s[k]
    per_doc.setdefault(k[0], []).append(d)
    if d < -0.001:
        reg += 1
        loss += -d
        print(f"  REGRESS {k[0]} p{k[1]}: {off_s[k]:.4f} -> {on_s[k]:.4f} ({d:+.4f})")
    elif d > 0.001:
        imp += 1
        gain += d
        print(f"  improve {k[0]} p{k[1]}: {off_s[k]:.4f} -> {on_s[k]:.4f} ({d:+.4f})")

print("\n=== per-doc mean delta (ON - OFF), |Δ|>0.0005 ===")
for dd, ds in sorted(per_doc.items()):
    m = sum(ds) / len(ds)
    if abs(m) > 0.0005:
        print(f"  {dd}: meanΔ={m:+.4f}  ({len(ds)} pages)")
print(f"\nTOTAL improved={imp} regressed={reg}  gain={gain:.4f} loss={loss:.4f} net={gain-loss:+.4f}")
