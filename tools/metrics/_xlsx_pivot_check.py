# -*- coding: utf-8 -*-
"""Hold Oxi's computed pivot grids (`oxi-xlsx-renderer <book> --pivot`) against
Excel's own refreshed pivot (<name>.truth.json from _xlsx_pivot_gen.py).

    python tools/metrics/_xlsx_pivot_check.py [dir]
"""
import json, os, subprocess, sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
RENDERER = Path(os.environ.get("OXI_XLSX_RENDERER", REPO / "tools/oxi-xlsx-renderer/target/release/oxi-xlsx-renderer.exe"))


def norm(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, float):
        return round(v, 9)
    return v


def main():
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else REPO / "pipeline_data" / "xlsx_pivot"
    ok = 0
    books = sorted(root.glob("*.xlsx"))
    for book in books:
        truth = json.loads(book.with_suffix(".truth.json").read_text(encoding="cp932"))
        run = subprocess.run([str(RENDERER), str(book), "--pivot"], capture_output=True, text=True, encoding="utf-8")
        got = json.loads(run.stdout or "[]")
        if not got:
            print(f"  {book.stem:28s} NO PIVOT READ  {run.stderr[:100]}")
            continue
        g = got[0]
        want = [[norm(v) for v in row] for row in truth["grid"]]
        mine = [[norm(v) for v in row] for row in g["grid"]]
        # Excel's used range is rectangular; pad ours to compare.
        width = max([len(r) for r in want + mine] or [0])
        want = [r + [None] * (width - len(r)) for r in want]
        mine = [r + [None] * (width - len(r)) for r in mine]
        origin_ok = (g["origin"][0] + 1, g["origin"][1] + 1) == (truth["first_row"], truth["first_col"])
        same = want == mine and origin_ok
        ok += same
        if same:
            print(f"  {book.stem:28s} OK")
        else:
            print(f"  {book.stem:28s} DIFF origin {'ok' if origin_ok else g['origin']}")
            for i in range(max(len(want), len(mine))):
                a = want[i] if i < len(want) else None
                b = mine[i] if i < len(mine) else None
                if a != b:
                    print(f"      row {i}: excel {a}\n              oxi   {b}")
    print(f"{ok}/{len(books)} pivots match Excel cell for cell")


if __name__ == "__main__":
    main()
