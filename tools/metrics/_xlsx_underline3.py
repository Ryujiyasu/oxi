"""Is the hyperlink's rule low, or is the whole line low?

`_xlsx_underline2.py` says Excel puts the rule exactly where the face declares
it, on every face and size measured -- so a rule that lands a pixel low in
`dendeba_kmc` cannot be the rule's own metric. Either the LINE is a pixel low
and the rule merely makes it visible, or it really is the rule.

The hyperlink is blue, which separates it from every border and every black
run on the sheet. This reads the blue ink out of the Excel picture and out of
Oxi's, band by band, and prints for each band where its letters start and
where its rule sits. If both move together the line is at fault; if the
letters agree and only the rule moves, the rule is.

Run: python tools/metrics/_xlsx_underline3.py <stem>
"""

from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
DIFF = REPO / "pipeline_data" / "xlsx_diff_probe"


def blue(path):
    """Rows that carry blue ink, and how much of it, per row."""
    rgb = np.asarray(Image.open(path).convert("RGB")).astype(np.int16)
    mask = (rgb[:, :, 2] - rgb[:, :, 0] > 40) & (rgb[:, :, 2] > 90)
    return mask


def bands(mask, gap=6):
    """Contiguous groups of blue rows, split where the sheet goes quiet."""
    rows = [y for y in range(mask.shape[0]) if mask[y].sum() >= 3]
    if not rows:
        return []
    out, start, last = [], rows[0], rows[0]
    for y in rows[1:]:
        if y - last > gap:
            out.append((start, last))
            start = y
        last = y
    out.append((start, last))
    return out


def rule_of(mask, top, bottom):
    """The widest row of the band, which for an underlined run is its rule."""
    counts = [(y, int(mask[y].sum())) for y in range(top, bottom + 1)]
    widest = max(counts, key=lambda held: held[1])
    span = int(mask[top:bottom + 1].any(axis=0).sum())
    return widest, span, counts


def main(stem):
    excel = DIFF / f"{stem}.excel.png"
    oxi = DIFF / f"{stem}.oxi.png"
    if not (excel.exists() and oxi.exists()):
        print(f"missing {excel.name} / {oxi.name} -- run xlsx_pixel_diff.py first")
        return 1
    xl, ox = blue(excel), blue(oxi)
    print(f"  blue ink: Excel {int(xl.sum())}px, Oxi {int(ox.sum())}px\n")
    xb, ob = bands(xl), bands(ox)
    print(f"  {len(xb)} blue band(s) in Excel, {len(ob)} in Oxi\n")
    print(f"  {'band':>4}  {'Excel top':>9} {'rule':>5} {'w':>4}  "
          f"{'Oxi top':>8} {'rule':>5} {'w':>4}   {'d top':>5} {'d rule':>6}")
    for i, ((xt, xbm), (ot, obm)) in enumerate(zip(xb, ob)):
        (xr, xw), xs, _ = rule_of(xl, xt, xbm)
        (orr, ow), os_, _ = rule_of(ox, ot, obm)
        print(f"  {i:>4}  {xt:>9} {xr:>5} {xw:>4}  {ot:>8} {orr:>5} {ow:>4}"
              f"   {ot - xt:>+5} {orr - xr:>+6}")
    if len(xb) != len(ob):
        print("\n  band counts differ -- the two are not the same lines,"
              " so the columns above are not paired")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main(sys.argv[1] if len(sys.argv) > 1
                          else "5c74ec72c6e1_h2daa2023_dendeba_kmc"))
