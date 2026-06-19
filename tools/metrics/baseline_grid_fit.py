# -*- coding: utf-8 -*-
"""Systematic reverse-engineering of Word's baseline GRID (device-snap model).

Session-1 foundation for the position device-snap multi-session effort. For a docx,
extracts Word's per-line baseline Y (via the cached *_word.json from word_pdf_glyphs.py,
or exports it), groups by font size, and FITS the device-snap model:

    baseline[i] = anchor + snap( i * base , delta )          (per font size)

minimizing the residual over (base, delta, anchor-phase). Answers the open question:
is the ~25mpt residual from the per-size BASE (83/64 deviation) or the snap PHASE?

Outputs, per (doc, font_size): best base (vs 83/64*fs), best delta, best phase,
and the residual at the best fit AND at the canonical (base=83/64, delta=0.12).

Usage: python baseline_grid_fit.py <word.json> [word2.json ...]
       (each word.json from tools/metrics/word_pdf_glyphs.py)
cp932-safe: no Japanese in code.
"""
import json, sys, os
import numpy as np
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

EI = 83.0 / 64.0  # 1.296875


def baselines_by_size(word_json, page=None):
    """Return {fs: [sorted baseline y]} for body lines, per page (or all)."""
    pages = json.load(open(word_json, encoding="utf-8"))["pages"]
    if page is not None:
        pages = [pages[page]]
    from collections import defaultdict
    out = defaultdict(list)
    for pg in pages:
        gs = sorted(pg["glyphs"], key=lambda g: (g["y"], g["x"]))
        lines, cur = [], []
        for g in gs:
            if cur and g["y"] - cur[-1]["y"] > 4.0:
                lines.append(cur); cur = []
            cur.append(g)
        if cur:
            lines.append(cur)
        for l in lines:
            fs = round(float(np.median([g["fs"] for g in l])), 1)
            out[fs].append(float(np.median([g["y"] for g in l])))
    return out


def fit_size(ys, fs):
    """Fit baseline[i]=anchor+snap(i*base, delta) to the run of consecutive body
    lines of one size. Returns dict of best params + residuals."""
    ys = np.array(sorted(ys))
    # restrict to consecutive body lines (delta within 1.5x of size*EI)
    base0 = fs * EI
    dys = np.diff(ys)
    keep = (dys > base0 * 0.6) & (dys < base0 * 1.5)
    if keep.sum() < 4:
        return None
    # take the longest consecutive run
    idx = np.where(keep)[0]
    # build cumulative from the first kept line
    runs = []
    start = idx[0]
    prev = idx[0]
    for k in idx[1:]:
        if k == prev + 1:
            prev = k
        else:
            runs.append((start, prev)); start = k; prev = k
    runs.append((start, prev))
    s, e = max(runs, key=lambda r: r[1] - r[0])
    seg = ys[s:e + 2]  # baselines of the run
    n = len(seg)
    if n < 4:
        return None

    def residual(base, delta):
        # predicted positions: best anchor minimizes |seg - (anchor + snap(j*base, delta))|
        j = np.arange(n)
        snapped = np.round(j * base / delta) * delta  # relative grid from line 0
        # best constant anchor = mean(seg - snapped)
        anchor = np.mean(seg - snapped)
        pred = anchor + snapped
        return float(np.sqrt(np.mean((seg - pred) ** 2))), anchor

    # canonical model
    can_res, _ = residual(base0, 0.12)
    # sweep base around 83/64 and delta
    best = (1e9, base0, 0.12)
    for base in np.arange(base0 - 0.08, base0 + 0.08, 0.004):
        for delta in [0.06, 0.08, 0.10, 0.12, 0.15, 0.20, 0.24]:
            r, _ = residual(base, delta)
            if r < best[0]:
                best = (r, base, delta)
    return {
        "fs": fs, "n": n,
        "mean_delta": float(np.mean(np.diff(seg))),
        "base_8364": base0,
        "canonical_res_mpt": can_res * 1000,
        "best_res_mpt": best[0] * 1000,
        "best_base": best[1], "best_base_factor": best[1] / fs,
        "best_delta": best[2],
    }


def main():
    print(f"{'doc':18} {'fs':>5} {'n':>3} {'mean_d':>7} {'8364*fs':>8} "
          f"{'canon_res':>9} {'best_res':>8} {'best_base':>9} {'b/fs':>7} {'best_d':>6}")
    for wj in sys.argv[1:]:
        name = os.path.basename(wj).replace("_word.json", "")
        bysz = baselines_by_size(wj)
        for fs in sorted(bysz):
            if len(bysz[fs]) < 5:
                continue
            r = fit_size(bysz[fs], fs)
            if not r:
                continue
            print(f"{name:18} {r['fs']:5.1f} {r['n']:3d} {r['mean_delta']:7.3f} "
                  f"{r['base_8364']:8.3f} {r['canonical_res_mpt']:8.1f}m {r['best_res_mpt']:7.1f}m "
                  f"{r['best_base']:9.3f} {r['best_base_factor']:7.4f} {r['best_delta']:6.2f}")


if __name__ == "__main__":
    main()
