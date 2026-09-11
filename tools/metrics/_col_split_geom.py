# -*- coding: utf-8 -*-
"""Where exactly does Word put every paragraph of a two-column probe?

The split can be read from the x positions alone, but not WHY Word chose it.
That needs the geometry: how tall each column ended up, and where the
section-ending paragraph mark sits. Word reports both, one paragraph at a time.

    python tools/metrics/_col_split_geom.py pos_8lines_tall24hp_at0_compat15
"""
import sys
from pathlib import Path

import win32com.client

PROBES = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "column_split"


def main() -> int:
    names = sys.argv[1:] or ["pos_8lines_tall24hp_at0_compat15"]
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        for name in names:
            path = PROBES / f"{name}.docx"
            doc = app.Documents.Open(str(path), False, True)
            print(f"\n{name}")
            rows = []
            for n, para in enumerate(doc.Paragraphs):
                rng = doc.Range(para.Range.Start, para.Range.Start)
                # 5 = horizontal, 6 = vertical, both relative to the page
                x = round(float(rng.Information(5)), 2)
                y = round(float(rng.Information(6)), 2)
                text = para.Range.Text.replace("\r", "")
                rows.append((n, x, y, text))
                print(f"  {n:2}  x={x:7.2f}  y={y:7.2f}  {text!r}")
            left = min(r[1] for r in rows)
            for side in (0, 1):
                col = [r for r in rows if (abs(r[1] - left) < 1.0) == (side == 0)]
                if len(col) > 1:
                    print(f"  column {side}: {len(col)} paragraphs, "
                          f"y {col[0][2]:.2f} to {col[-1][2]:.2f}, "
                          f"pitch {[round(b[2]-a[2], 2) for a, b in zip(col, col[1:])]}")
            doc.Close(False)
    finally:
        app.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
