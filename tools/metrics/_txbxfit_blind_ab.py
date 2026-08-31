# -*- coding: utf-8 -*-
"""S1266 A/B on the JA blind sets (measurement target, never a fix target).

Same method as tools/metrics/ssim_ab.py but against the blind sets' cached
Word ExportAsFixedFormat PDFs: render each doc twice with the SAME binary
(A = OXI_S1266_DISABLE=1, B = default), score both against the Word PDF page
by page, and report per-doc common/penalized means both ways.  Rendering both
arms now means no cached-PNG staleness can leak into the number.

Usage: python tools/metrics/_txbxfit_blind_ab.py [doc-prefix ...]
"""
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
BENCH = REPO / "pipeline_data" / "ja_benchmark"
DWRITE = os.environ.get("OXI_DWRITE_EXE") or str(
    REPO / "tools" / "oxi-dwrite-renderer" / "target" / "release" / "oxi-dwrite-renderer.exe"
)
DPI = 150
# Comma-separated names gate a BUNDLE: arm A sets them all, arm B clears them
# all -- the only honest way to measure a pair whose members compensate each
# other (S1266 restores the text, S1267 puts the box where Word puts it).
FLAGS = [f for f in os.environ.get("OXI_AB_FLAG", "OXI_S1266_DISABLE").split(",") if f]
SETS = ("blind50", "blindB50", "blindC50")


def selections(setname):
    key = {"blind50": "_final_jablind50.json",
           "blindB50": "_final_jablindB50.json",
           "blindC50": "_final_jablindC50.json"}[setname]
    data = json.loads((BENCH / key).read_text(encoding="utf-8"))
    rows = []
    for kind, docs in data.items():
        for doc in docs:
            path = Path(doc["path"])
            rows.append((f"{kind}__{path.stem}", path))
    return rows


def rgb_png(path):
    with Image.open(path) as im:
        return np.asarray(im.convert("RGB"))


def rgb_pdf(pdf, i):
    pix = pdf[i].get_pixmap(matrix=fitz.Matrix(DPI / 72, DPI / 72), alpha=False)
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]


def resize(cand, ref):
    if cand.shape == ref.shape:
        return cand
    return np.asarray(Image.fromarray(cand).resize((ref.shape[1], ref.shape[0]),
                                                   Image.Resampling.LANCZOS))


def score(ref, cand):
    return float(structural_similarity(ref, resize(cand, ref), channel_axis=2, data_range=255))


def render(docx, outdir, arm_a):
    env = dict(os.environ)
    for f in FLAGS:
        if arm_a:
            env[f] = "1"
        else:
            env.pop(f, None)
    Path(outdir).mkdir(parents=True, exist_ok=True)
    subprocess.run([DWRITE, str(docx), str(Path(outdir) / "p"), str(DPI)],
                   capture_output=True, timeout=1800, env=env)
    pages, i = [], 1
    while (Path(outdir) / ("p_p%d.png" % i)).is_file():
        pages.append(Path(outdir) / ("p_p%d.png" % i))
        i += 1
    return pages


def main():
    filt = sys.argv[1:]
    rows = []
    for s in SETS:
        for doc, path in selections(s):
            if filt and not any(doc.startswith(f) for f in filt):
                continue
            rows.append((s, doc, path))
    print("%d docs | flags %s (A=disabled, B=default)" % (len(rows), ",".join(FLAGS)))
    changed = []
    with tempfile.TemporaryDirectory(prefix="oxi_blind_ab_") as tmp:
        for setname, doc, path in rows:
            wp = BENCH / ("ssim_" + setname) / "word_pdf" / (doc + ".pdf")
            if not wp.is_file():
                continue
            pa = render(path, Path(tmp) / "A" / doc, True)
            pb = render(path, Path(tmp) / "B" / doc, False)
            same = len(pa) == len(pb) and all(
                a.read_bytes() == b.read_bytes() for a, b in zip(pa, pb))
            if same:
                continue
            word = fitz.open(wp)
            n = word.page_count
            sa = sb = 0.0
            for i in range(n):
                ref = rgb_pdf(word, i)
                if i < len(pa):
                    sa += score(ref, rgb_png(pa[i]))
                if i < len(pb):
                    sb += score(ref, rgb_png(pb[i]))
            word.close()
            pen_a = sa / max(n, len(pa))
            pen_b = sb / max(n, len(pb))
            com_a = sa / min(n, len(pa)) if pa else 0.0
            com_b = sb / min(n, len(pb)) if pb else 0.0
            changed.append((setname, doc, n, len(pa), len(pb), com_a, com_b, pen_a, pen_b))
            print("  %-10s %-40s wordpg=%d pg %d->%d  common %.4f->%.4f (%+.4f)  penalized %.4f->%.4f (%+.4f)"
                  % (setname, doc, n, len(pa), len(pb), com_a, com_b, com_b - com_a,
                     pen_a, pen_b, pen_b - pen_a), flush=True)
    print()
    if not changed:
        print("no doc changed bytes")
        return
    dcom = sum(r[6] - r[5] for r in changed)
    dpen = sum(r[8] - r[7] for r in changed)
    up = sum(1 for r in changed if r[8] - r[7] > 0.0005)
    down = sum(1 for r in changed if r[8] - r[7] < -0.0005)
    print("changed %d docs | sum d(common)=%+.4f sum d(penalized)=%+.4f | improved %d regressed %d"
          % (len(changed), dcom, dpen, up, down))
    print("blind-set totals move by d(penalized)/50 per set; see per-doc rows above.")


if __name__ == "__main__":
    main()
