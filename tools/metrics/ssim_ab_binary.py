# -*- coding: utf-8 -*-
"""SSIM A/B between two BUILDS of the renderer, on the corpus.

`ssim_ab.py` switches an environment flag, which is the right instrument for a
rule that ships behind one. A rule that REPLACES another cannot be flagged —
there is no build in which both exist — so the only honest before-and-after is
two binaries rendering the same documents against the same Word references.

    python tools/metrics/ssim_ab_binary.py <old-renderer.exe> [base-prefix...]

Reports net (new - old). Same SSIM as the production gate.
"""
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

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
NEW = _REPO / "tools" / "oxi-dwrite-renderer" / "target" / "release" / "oxi-dwrite-renderer.exe"
DOCS = _REPO / "tools" / "golden-test" / "documents" / "docx"


def score(word_png: str, oxi_png: str) -> float:
    w = _load_rgb(word_png)
    o = _resize_to_match(_load_rgb(oxi_png), w)
    return ssim(w, o, full=False, channel_axis=2, data_range=255)


def render(exe: Path, docx: str, outdir: Path) -> list:
    outdir.mkdir(parents=True, exist_ok=True)
    subprocess.run([str(exe), docx, str(outdir / "p"), str(RENDER_DPI)],
                   capture_output=True, timeout=300)
    pages, i = [], 1
    while (outdir / f"p_p{i}.png").exists():
        pages.append(str(outdir / f"p_p{i}.png"))
        i += 1
    return pages


def find(base: str):
    exact = DOCS / (base + ".docx")
    if exact.exists():
        return str(exact.resolve())
    near = sorted(p for p in DOCS.glob(base.split("_")[0] + "*.docx")
                  if not p.name.startswith("~$"))
    return str(near[0].resolve()) if near else None


def main() -> int:
    if len(sys.argv) < 2:
        print(f"usage: {Path(sys.argv[0]).name} <old-renderer.exe> [base-prefix...]")
        return 2
    old = Path(sys.argv[1])
    if not old.is_file():
        print(f"no renderer at {old}")
        return 1
    wanted = sys.argv[2:]
    bases = sorted({re.sub(r"_p\d+$", "", n) for n in os.listdir(WORD_PNG_DIR)})
    if wanted:
        bases = [b for b in bases if any(b.startswith(w) for w in wanted)]

    changed, checked, seen = [], 0, set()
    with tempfile.TemporaryDirectory() as tmp:
        for base in bases:
            docx = find(base)
            if not docx or docx in seen:
                continue
            seen.add(docx)
            checked += 1
            a = Path(tmp) / "old" / Path(docx).stem
            b = Path(tmp) / "new" / Path(docx).stem
            pa, pb = render(old, docx, a), render(NEW, docx, b)
            differs = len(pa) != len(pb) or any(
                open(x, "rb").read() != open(y, "rb").read() for x, y in zip(pa, pb))
            if differs:
                changed.append((base, a, b, len(pa), len(pb)))
        print(f"checked {checked}; {len(changed)} render differently")

        total, wins, losses, rows = 0.0, 0, 0, []
        for base, a, b, na, nb in changed:
            wdir = Path(WORD_PNG_DIR) / base
            net, pages, i = 0.0, 0, 1
            while True:
                wp = wdir / f"page_{i:04d}.png"
                ap, bp = a / f"p_p{i}.png", b / f"p_p{i}.png"
                if not wp.exists() or not ap.exists() or not bp.exists():
                    break
                try:
                    net += score(str(wp), str(bp)) - score(str(wp), str(ap))
                    pages += 1
                except Exception:
                    pass
                i += 1
            if not pages:
                rows.append((base, None, na, nb))
                continue
            total += net
            wins += net > 0.0005
            losses += net < -0.0005
            rows.append((base, net, na, nb))

        for base, net, na, nb in sorted(rows, key=lambda r: r[1] if r[1] is not None else 0):
            if net is None:
                print(f"  {base}: no Word reference pages (pages {na}/{nb})")
            else:
                mark = " <<< REGRESS" if net < -0.0005 else (" >>> improve" if net > 0.0005 else "")
                print(f"  {base}: pages={na}/{nb} net(new-old)={net:+.4f}{mark}")
        print(f"\nTOTAL net(new-old)={total:+.4f}; improved {wins}, regressed {losses}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
