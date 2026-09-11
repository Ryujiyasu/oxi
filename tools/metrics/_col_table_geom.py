# -*- coding: utf-8 -*-
"""Exact per-line geometry for the table-in-a-column probes, saved once.

`_col_split_lines.py` reads Word's page rectangles, which is the only way to
see a split inside a paragraph but reports whole points. Deriving a rule needs
the sub-point truth, and that comes from Information(6), one paragraph at a
time. A table ROW is one line of the column however many cells sit on it, so
cells sharing a vertical position collapse to one entry here.

    python tools/metrics/_col_table_geom.py [out.json]

Writes every probe's two columns as lists of (top, kind), so a candidate rule
can be tried against all of them without opening Word again.
"""
import json
import sys
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
PROBES = REPO / "tests" / "fixtures" / "column_split"
MIDDLE = 300.0  # the page's centre; the two column bands sit either side


def measure(app, path: Path) -> dict:
    doc = app.Documents.Open(str(path), False, True)
    try:
        seen = {}
        for para in doc.Paragraphs:
            rng = doc.Range(para.Range.Start, para.Range.Start)
            x = float(rng.Information(5))
            y = round(float(rng.Information(6)), 2)
            text = para.Range.Text.replace("\r", "").replace("\x07", "")
            in_cell = bool(para.Range.Information(12))  # wdWithInTable
            side = "right" if x >= MIDDLE else "left"
            # Cells of one row share a top; keep the row once, and keep the
            # leftmost x so a cell cannot be mistaken for its own column.
            key = (side, y)
            if key not in seen or x < seen[key][0]:
                seen[key] = (x, "cell" if in_cell else "text", text.strip()[:12])
        out = {"left": [], "right": []}
        for (side, y), (_x, kind, text) in sorted(seen.items(), key=lambda kv: (kv[0][0], kv[0][1])):
            out[side].append({"top": y, "kind": kind, "text": text})
        return out
    finally:
        doc.Close(False)


def main() -> int:
    where = Path(sys.argv[1]) if len(sys.argv) > 1 else \
        REPO / "pipeline_data" / "col_table_geom.json"
    probes = sorted(PROBES.glob("tab_*rows_at*_compat15.docx"))
    if not probes:
        print("no table probes; run tools/metrics/_col_table_gen.py first")
        return 1
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = False
    found = {}
    try:
        for path in probes:
            found[path.stem] = measure(app, path)
            got = found[path.stem]
            print(f"  {path.stem:30s} left {len(got['left']):2} right {len(got['right']):2}",
                  flush=True)
    finally:
        app.Quit()
    where.parent.mkdir(parents=True, exist_ok=True)
    where.write_text(json.dumps(found, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"written: {where}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
