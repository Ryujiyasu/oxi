# -*- coding: utf-8 -*-
"""Ask Word where it broke each two-column probe, and the engine the same.

Word does not report a column index, but it reports each paragraph's horizontal
position on the page, and in a two-column section that is the column. So the
split is read as "how many paragraphs sit in the left column".

    python tools/metrics/_col_split_word.py tests/fixtures/column_split

Prints one row per probe: what Word did, what the engine did, and whether they
agree. The point is the boundary — the last-line size at which Word moves from
five left rows to six — not any single row.
"""
import sys
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
GDI = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"


def word_split(app, path: Path) -> tuple[int, int] | None:
    """(paragraphs in the left column, total) as Word lays it out."""
    doc = app.Documents.Open(str(path.resolve()), False, True)
    try:
        xs = []
        for para in doc.Paragraphs:
            rng = doc.Range(para.Range.Start, para.Range.Start)
            # 5 = wdHorizontalPositionRelativeToPage
            xs.append(round(float(rng.Information(5)), 1))
        if not xs:
            return None
        left = min(xs)
        return sum(1 for x in xs if abs(x - left) < 1.0), len(xs)
    finally:
        doc.Close(False)


def oxi_split(path: Path) -> tuple[int, int] | None:
    import json
    import subprocess
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "layout.json"
        subprocess.run([str(GDI), str(path.resolve()), str(Path(tmp) / "p"),
                        "150", f"--dump-layout={dump}"],
                       capture_output=True)
        if not dump.is_file():
            return None
        data = json.loads(dump.read_text(encoding="utf-8"))
    xs = []
    for page in data.get("pages", {}).values() if isinstance(data.get("pages"), dict) \
            else data.get("pages", []):
        for rec in (page if isinstance(page, list) else page.get("records", [])):
            if rec.get("text"):
                xs.append(round(float(rec.get("x", 0)), 1))
    if not xs:
        return None
    left = min(xs)
    return sum(1 for x in xs if abs(x - left) < 1.0), len(xs)


def main() -> int:
    where = Path(sys.argv[1] if len(sys.argv) > 1 else "tests/fixtures/column_split")
    probes = sorted(where.glob("*.docx"))
    if not probes:
        print(f"nothing to measure in {where}")
        return 1

    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = False
    rows = []
    try:
        for path in probes:
            got = word_split(app, path)
            mine = oxi_split(path)
            rows.append((path.stem, got, mine))
            agree = "" if not got or not mine else ("ok" if got[0] == mine[0] else "DIFFERS")
            print("  %-38s word %-8s oxi %-8s %s"
                  % (path.stem[:38],
                     f"{got[0]}/{got[1]}" if got else "-",
                     f"{mine[0]}/{mine[1]}" if mine else "-",
                     agree), flush=True)
    finally:
        app.Quit()

    print("\nwhere Word moves the split, per line count and compat mode:")
    seen = {}
    for name, got, _mine in rows:
        if not got:
            continue
        bits = name.split("_")
        count = int(bits[1].replace("lines", ""))
        half = int(bits[2].replace("last", "").replace("hp", ""))
        compat = int(bits[3].replace("compat", ""))
        seen.setdefault((count, compat), []).append((half, got[0]))
    for key in sorted(seen):
        pairs = sorted(seen[key])
        line = "  ".join(f"{h/2:g}pt:{left}" for h, left in pairs)
        print(f"  {key[0]} lines, compat {key[1]}:  {line}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
