"""S468 per-line glyph-ink diagnostic: why does VSNAP regress CJK but help Latin?
For a (doc,page), load Word PNG + Oxi OFF + Oxi ON (from the gate's TMP renders).
Compute each image's row ink profile (darkness per row), cross-correlate OFF-vs-Word
and ON-vs-Word over integer row shifts -> best global shift + peak correlation.
If ON's best-shift moves toward 0 and peak corr rises, VSNAP aligned better;
if ON shifts away / corr drops, VSNAP misaligned (the regression mechanism)."""
import os, sys
import numpy as np
from PIL import Image

REPO = r"C:\Users\ryuji\oxi-main"
WORD = os.path.join(REPO, "pipeline_data", "word_png")
TMP = r"C:\Users\ryuji\AppData\Local\Temp\s468vgate"


def row_profile(path, ref_shape=None):
    im = Image.open(path).convert("L")
    a = np.array(im)
    if ref_shape is not None and a.shape != ref_shape:
        im = im.resize((ref_shape[1], ref_shape[0]), Image.LANCZOS)
        a = np.array(im)
    ink = (255.0 - a.astype(np.float64))  # darkness
    return ink.sum(axis=1), a.shape


def best_shift(prof_ref, prof_test, maxshift=12):
    # normalize
    r = prof_ref - prof_ref.mean()
    t = prof_test - prof_test.mean()
    best = (None, -2.0)
    n = len(r)
    for s in range(-maxshift, maxshift + 1):
        if s >= 0:
            rr = r[s:]; tt = t[:n - s]
        else:
            rr = r[:n + s]; tt = t[-s:]
        if len(rr) < 50:
            continue
        denom = (np.linalg.norm(rr) * np.linalg.norm(tt))
        if denom == 0:
            continue
        c = float(np.dot(rr, tt) / denom)
        if c > best[1]:
            best = (s, c)
    return best


def analyze(doc_dir, doc_id, pages):
    wdir = os.path.join(WORD, doc_id)
    for pg in pages:
        wpng = os.path.join(wdir, "page_%04d.png" % pg)
        offp = os.path.join(TMP, "off_%s_p%d.png" % (doc_id, pg))
        onp = os.path.join(TMP, "on_%s_p%d.png" % (doc_id, pg))
        if not (os.path.exists(wpng) and os.path.exists(offp) and os.path.exists(onp)):
            print("  p%d: missing png" % pg); continue
        wprof, wshape = row_profile(wpng)
        offprof, _ = row_profile(offp, wshape)
        onprof, _ = row_profile(onp, wshape)
        soff, coff = best_shift(wprof, offprof)
        son, con = best_shift(wprof, onprof)
        print("  p%-2d  OFF: best_shift=%+d corr=%.4f | ON: best_shift=%+d corr=%.4f | dcorr=%+.4f"
              % (pg, soff, coff, son, con, con - coff))


def main():
    targets = [
        ("0e7af1ae8f21_20230331_resources_open_data_contract_sample_00", [2, 4, 6]),
        ("c7b923e5c616_20240705_resources_data_outline_08", [1, 2, 3]),
        ("gen2_051_Security_Policy", [1, 2]),
        ("gen2_054_Audit_Report", [1, 2]),
        ("34140b9c5662_index-14", [1, 5]),
    ]
    for doc_id, pages in targets:
        print("=== %s ===" % doc_id)
        analyze(WORD, doc_id, pages)


if __name__ == "__main__":
    main()
