# -*- coding: utf-8 -*-
"""SSIM A/B for a document that is NOT in the word_png sentinel corpus.

kojin is the document S1151 was derived on and it has no cached Word reference,
so ssim_ab.py cannot see it.  Rasterise Word's own PDF (already cached by
_kojin_rowgeom.py) at the pipeline DPI and score both Oxi variants against it.
The absolute number is not comparable to the sentinel (PDF raster, not EMF), but
the A/B delta is, because both sides use the same reference.

    python _kojin_ssim_ab.py OXI_S1151=1        # A = var set, B = default
    OXI_DOC=b837 python _kojin_ssim_ab.py OXI_S1151=1
"""
import os
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(HERE))
from pipeline.config import RENDER_DPI  # noqa: E402
from pipeline.ssim_calculator import _load_rgb, _resize_to_match  # noqa: E402
from skimage.metrics import structural_similarity as ssim  # noqa: E402

import _kojin_rowgeom as K  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
DW = REPO / "tools" / "oxi-dwrite-renderer" / "target" / "release" / "oxi-dwrite-renderer.exe"
ARG = sys.argv[1] if len(sys.argv) > 1 else "OXI_S1151=1"
AENV = []
for part in ARG.split(","):
    if not part:
        continue
    k, _, v = part.partition("=")
    AENV.append((k, v or "1"))


def word_pages(outdir):
    import fitz
    doc = fitz.open(K._ensure_pdf())
    outdir.mkdir(parents=True, exist_ok=True)
    out = []
    for i in range(doc.page_count):
        p = outdir / ("page_%04d.png" % (i + 1))
        if not p.exists():
            doc[i].get_pixmap(dpi=RENDER_DPI).save(str(p))
        out.append(p)
    return out


def oxi_pages(outdir, on):
    env = dict(os.environ)
    for k, v in AENV:
        if on:
            env[k] = v
        else:
            env.pop(k, None)
    outdir.mkdir(parents=True, exist_ok=True)
    subprocess.run([str(DW), K.DOCX, str(outdir / "p"), str(RENDER_DPI)],
                   check=True, capture_output=True, env=env)
    # NUMERIC sort -- plain sorted() gives p1, p10, p11, ... p2 and silently
    # scores every page against the wrong reference.
    import re
    return sorted(outdir.glob("p*.png"),
                  key=lambda p: int(re.search(r"(\d+)(?=\.png$)", p.name).group(1)))


def main():
    tmp = Path(tempfile.gettempdir()) / ("ssimab_" + K.DOC)
    wp = word_pages(tmp / "word")
    a = oxi_pages(tmp / "a", True)
    b = oxi_pages(tmp / "b", False)
    print("%-6s %-9s %-9s %s" % ("page", "A(" + AENV[0][0] + ")", "B(default)", "delta"))
    sa = sb = 0.0
    n = min(len(wp), len(a), len(b))
    for i in range(n):
        w = _load_rgb(str(wp[i]))
        va = ssim(w, _resize_to_match(_load_rgb(str(a[i])), w), channel_axis=2,
                  data_range=255)
        vb = ssim(w, _resize_to_match(_load_rgb(str(b[i])), w), channel_axis=2,
                  data_range=255)
        sa += va
        sb += vb
        flag = "" if abs(va - vb) < 1e-4 else ("  <-- A better" if va > vb else "  <-- B better")
        print("%-6d %-9.4f %-9.4f %+.4f%s" % (i + 1, va, vb, va - vb, flag))
    print("pages word=%d A=%d B=%d" % (len(wp), len(a), len(b)))
    print("MEAN   %-9.4f %-9.4f %+.4f" % (sa / n, sb / n, (sa - sb) / n))


if __name__ == "__main__":
    main()
