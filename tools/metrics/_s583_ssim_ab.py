# -*- coding: utf-8 -*-
"""S583 — SSIM sentinel for the no-type-docGrid page-bottom ink-leniency.

Render every word_png doc base with DWrite A (OXI_S583_DISABLE=1) and B
(default ON). First compare PNG bytes A vs B; S581 only fires on no-type
docGrid docs with a page-bottom line in the [ink, natural_lh] window, so the
vast majority are byte-identical and skipped. For the docs that DO differ,
SSIM each page vs the cached word_png both ways and report ON-OFF.

Usage: python _s576_ssim_ab.py            (all word_png bases)
       python _s576_ssim_ab.py <base...>  (specific bases)
"""
import os
import re
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

# word_png bases (strip _pN)
bases = sorted({re.sub(r"_p\d+$", "", n) for n in os.listdir(WORD_PNG_DIR)})
args = sys.argv[1:]
if args:
    bases = [b for b in bases if any(b.startswith(a) or b == a for a in args)]

# map base -> docx: exact stem match first, then doc_id-token prefix.
def find_docx(base):
    exact = Path(DOCS_DIR) / (base + ".docx")
    if exact.exists():
        return os.path.abspath(str(exact))
    tok = base.split("_")[0]
    cands = sorted(p for p in Path(DOCS_DIR).glob(tok + "*.docx") if not p.name.startswith("~$"))
    return os.path.abspath(str(cands[0])) if cands else None


def render(docx, disable, outdir):
    env = dict(os.environ)
    if disable:
        env["OXI_S583_DISABLE"] = "1"
    else:
        env.pop("OXI_S583_DISABLE", None)
    Path(outdir).mkdir(parents=True, exist_ok=True)
    subprocess.run([DW, docx, str(Path(outdir) / "p"), str(RENDER_DPI)],
                   capture_output=True, timeout=300, env=env)
    pages = []
    i = 1
    while (Path(outdir) / f"p_p{i}.png").exists():
        pages.append(str(Path(outdir) / f"p_p{i}.png"))
        i += 1
    return pages


changed = []
checked = 0
with tempfile.TemporaryDirectory() as tmp:
    seen_docx = {}
    for base in bases:
        docx = find_docx(base)
        if not docx:
            continue
        if docx in seen_docx:
            continue  # render each docx once
        seen_docx[docx] = base
        checked += 1
        a_dir = Path(tmp) / "A" / Path(docx).stem
        b_dir = Path(tmp) / "B" / Path(docx).stem
        pa = render(docx, True, a_dir)
        pb = render(docx, False, b_dir)
        # byte compare
        diff = (len(pa) != len(pb))
        if not diff:
            for x, y in zip(pa, pb):
                if open(x, "rb").read() != open(y, "rb").read():
                    diff = True
                    break
        if diff:
            changed.append((base, docx, a_dir, b_dir, len(pa), len(pb)))

print(f"checked {checked} docx; {len(changed)} changed bytes A vs B")
for base, docx, a_dir, b_dir, na, nb in changed:
    print(f"  CHANGED {base}  pages A={na} B={nb}")

# SSIM only the changed docs
if changed:
    print("\n=== SSIM A(OFF) vs B(ON) for changed docs ===")
    for base, docx, a_dir, b_dir, na, nb in changed:
        # word_png pages: word_png/<base>/page_NNNN.png
        wdir = Path(WORD_PNG_DIR) / base
        i = 1
        net = 0.0
        rows = []
        while True:
            wp = wdir / f"page_{i:04d}.png"
            if not wp.exists():
                break
            a_png = Path(str(a_dir)) / f"p_p{i}.png"
            b_png = Path(str(b_dir)) / f"p_p{i}.png"
            if not a_png.exists() or not b_png.exists():
                break
            try:
                sa = calculate_ssim(str(wp), str(a_png))
                sb = calculate_ssim(str(wp), str(b_png))
            except Exception as e:
                rows.append((i, "err", str(e)[:40]))
                i += 1
                continue
            net += (sb - sa)
            rows.append((i, sa, sb))
            i += 1
        print(f"  {base}: net(ON-OFF)={net:+.4f}")
        for r in rows:
            if isinstance(r[1], float) and abs(r[2] - r[1]) > 1e-6:
                print(f"     p{r[0]}: OFF={r[1]:.4f} ON={r[2]:.4f} d={r[2]-r[1]:+.4f}")
