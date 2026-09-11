# -*- coding: utf-8 -*-
"""How far this engine is from Word on a two-column run holding a table.

The text-only probes agree 129/129 (`_col_split_sweep.py`). This is the family
that does not, and it is where a blind document's loss was traced to
(`OXI_PREFIX_TABLE_BALANCE_DISABLE` moved it 0.9154 -> 0.9607 and nothing else
did). It exists so the gap has a number and a target rather than a memory.

    python tools/metrics/_col_table_sweep.py

Word's answers are recorded — measured 2026-09-11 with `_col_split_lines.py`,
which walks Word's own page rectangles line by line. A table ROW is one line
of the column to Word, however many cells sit on it, so both sides count a
row once: distinct vertical positions within a column band.
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
GDI = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
PROBES = REPO / "tests" / "fixtures" / "column_split"

# (table rows, position among the 8 text lines) -> lines in the LEFT column,
# counting the one-column paragraph that follows the run.
WORD = {
    (1, 0): 5, (1, 2): 5, (1, 4): 7, (1, 6): 6, (1, 8): 6,
    (2, 0): 6, (2, 2): 6, (2, 4): 6, (2, 6): 7, (2, 8): 7,
    (3, 0): 6, (3, 2): 6, (3, 4): 7, (3, 6): 7, (3, 8): 7,
    (5, 0): 7, (5, 2): 8, (5, 4): 8, (5, 6): 8, (5, 8): 8,
}
MIDDLE = 300.0  # the page's centre, in points; the two column bands sit either side


def oxi_columns(path: Path) -> tuple[int, int] | None:
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "layout.json"
        subprocess.run([str(GDI), str(path), str(Path(tmp) / "p"), "150",
                        f"--dump-layout={dump}"], capture_output=True)
        if not dump.is_file():
            return None
        data = json.loads(dump.read_text(encoding="utf-8"))
    els = [e for page in data.get("pages", []) for e in page.get("elements", [])
           if e.get("type") == "text" and e.get("text")]
    if not els:
        return None
    left = {round(float(e["y"]), 1) for e in els if float(e["x"]) < MIDDLE}
    right = {round(float(e["y"]), 1) for e in els if float(e["x"]) >= MIDDLE}
    return len(left), len(right)


def main() -> int:
    agree = 0
    for (rows, at), want in sorted(WORD.items()):
        probe = PROBES / f"tab_{rows}rows_at{at}_compat15.docx"
        got = oxi_columns(probe) if probe.is_file() else None
        ok = got is not None and got[0] == want
        agree += ok
        if not ok:
            print(f"  table of {rows} rows at {at}: word left {want}  "
                  f"oxi left {got[0] if got else '-'} right {got[1] if got else '-'}")
    print(f"table in a column run: {agree}/{len(WORD)} agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
