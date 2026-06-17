# -*- coding: utf-8 -*-
"""S603/S604/S601 — SSIM sentinel for the combined ship:
  - S603: typed-grid page-bottom full-cell before a table
  - S604: body 約物 cap 3.0 -> 3.1
  - S601: line-end 約物 ぶら下げ default ON

Render every word_png doc base with DWrite A (OFF: OXI_S603_DISABLE=1
OXI_S601_DISABLE=1 OXI_S575_CAP=3.0) and B (default ON). Byte-compare A vs B;
the changes only affect typed-grid jc=both/legacy docs with 約物 line breaks or
a para-before-table page-bottom line, so most docs are byte-identical and
skipped. SSIM each changed doc's pages vs cached word_png both ways; report
net (ON-OFF) so a regression > a few thousandths surfaces.

Usage: python tools/metrics/_s603_ssim_ab.py            (all word_png bases)
       python tools/metrics/_s603_ssim_ab.py <base...>  (specific bases)
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

bases = sorted({re.sub(r"_p\d+$", "", n) for n in os.listdir(WORD_PNG_DIR)})
args = sys.argv[1:]
if args:
    bases = [b for b in bases if any(b.startswith(a) or b == a for a in args)]


def find_docx(base):
    exact = Path(DOCS_DIR) / (base + ".docx")
    if exact.exists():
        return os.path.abspath(str(exact))
    tok = base.split("_")[0]
    cands = sorted(p for p in Path(DOCS_DIR).glob(tok + "*.docx") if not p.name.startswith("~$"))
    return os.path.abspath(str(cands[0])) if cands else None


def render(docx, off, outdir):
    env = dict(os.environ)
    if off:
        env["OXI_S603_DISABLE"] = "1"
        env["OXI_S601_DISABLE"] = "1"
        env["OXI_S575_CAP"] = "3.0"
    else:
        env.pop("OXI_S603_DISABLE", None)
        env.pop("OXI_S601_DISABLE", None)
        env.pop("OXI_S575_CAP", None)
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
        if not docx or docx in seen_docx:
            continue
        seen_docx[docx] = base
        checked += 1
        a_dir = Path(tmp) / "A" / Path(docx).stem
        b_dir = Path(tmp) / "B" / Path(docx).stem
        pa = render(docx, True, a_dir)
        pb = render(docx, False, b_dir)
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

    if changed:
        print("\n=== SSIM A(OFF) vs B(ON) for changed docs ===")
        total_net = 0.0
        for base, docx, a_dir, b_dir, na, nb in changed:
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
            total_net += net
            print(f"  {base}: net(ON-OFF)={net:+.4f}")
            for r in rows:
                if isinstance(r[1], float) and abs(r[2] - r[1]) > 1e-6:
                    print(f"     p{r[0]}: OFF={r[1]:.4f} ON={r[2]:.4f} d={r[2]-r[1]:+.4f}")
        print(f"\nTOTAL net(ON-OFF) across changed docs = {total_net:+.4f}")
