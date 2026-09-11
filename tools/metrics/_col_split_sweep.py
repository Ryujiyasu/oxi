# -*- coding: utf-8 -*-
"""How often does the engine put the two-column break where Word puts it?

`_col_split_word.py` asks both sides, but it needs Word open, so re-running it
after every rebuild costs minutes and a COM session. Word's answers do not
change, so they are recorded here once and only the engine is re-measured.

    python tools/metrics/_col_split_sweep.py

The interesting family is the position sweep: the same tall line moved through
the run. It says what "tall" is measured against, which a probe that only ever
puts the tall line last cannot say.
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

# Measured with Word COM (`_col_split_word.py`, 2026-09-11). The count is
# paragraphs sitting at the page's left edge, which includes the one-column
# paragraph after the run — both sides are counted the same way, so the
# comparison holds without unpicking it.
WORD = {
    (8, 0): 5, (8, 1): 5, (8, 2): 5, (8, 3): 5, (8, 4): 5,
    (8, 5): 6, (8, 6): 6, (8, 7): 6,
    (10, 0): 6, (10, 1): 6, (10, 2): 6, (10, 3): 6, (10, 4): 6, (10, 5): 6,
    (10, 6): 7, (10, 7): 7, (10, 8): 7, (10, 9): 7,
    (12, 0): 7, (12, 1): 7, (12, 2): 7, (12, 3): 7, (12, 4): 7, (12, 5): 7,
    (12, 6): 7,
    (12, 7): 8, (12, 8): 8, (12, 9): 8, (12, 10): 8, (12, 11): 8,
}

# The other family: every row the same except the LAST, swept 10pt to 12pt.
# 20 half-points is the base, so `last: 20` is an evenly-sized run.
SIZES = {
    (6, 20): 4, (6, 21): 5, (6, 22): 5, (6, 23): 5, (6, 24): 5,
    (8, 20): 5, (8, 21): 6, (8, 22): 6, (8, 23): 6, (8, 24): 6,
    (10, 20): 6, (10, 21): 7, (10, 22): 7, (10, 23): 7, (10, 24): 7,
    (12, 20): 7, (12, 21): 8, (12, 22): 8, (12, 23): 8, (12, 24): 8,
    (14, 20): 8, (14, 21): 9, (14, 22): 9, (14, 23): 9, (14, 24): 9,
}


# A third family: the same even run, but with full-width lines above it on the
# same page, so the run does not start at the top. Word's answer never moves —
# it balances the RUN, not what is left of the page — and the count below is
# the whole left edge, prefix and trailing paragraph included.
# Paragraphs that WRAP, so the split can land inside one. Word does split
# mid-paragraph here (the right column starts mid-sentence), and the counts
# were read with `_col_split_lines.py`, which walks Word's own page rectangles
# line by line — paragraph positions cannot see a split inside a paragraph.
WRAPPED = {
    (2, 140): 4, (2, 200): 5,
    (3, 140): 6, (3, 200): 7,
    (4, 140): 7, (4, 200): 9,
    (5, 140): 9, (5, 200): 11,
}

# How a column run is CLOSED decides whether Word balances it at all.
# The count is the engine's: paragraphs with a text element at the left edge,
# so Word's own trailing empty paragraph does not appear in it.
CLOSING = {
    # the run is the document's last section, nothing closes it: Word FILLS
    "nothingafter": 8,
    # closed by a next-page section break: balanced (4 left, plus the
    # one-column paragraph that follows, which shares the left edge)
    "nextpage": 5,
    # closed by a continuous break whose section holds nothing: still balanced
    "emptyafter": 4,
}

PREFIXED = {
    (8, 1): 6, (8, 2): 7, (8, 3): 8, (8, 5): 10,
    (10, 1): 7, (10, 2): 8, (10, 3): 9, (10, 5): 11,
}


def oxi_left(path: Path) -> int | None:
    """Paragraphs the engine placed at the left edge."""
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "layout.json"
        subprocess.run([str(GDI), str(path), str(Path(tmp) / "p"), "150",
                        f"--dump-layout={dump}"], capture_output=True)
        if not dump.is_file():
            return None
        data = json.loads(dump.read_text(encoding="utf-8"))
    xs = [round(float(e["x"]), 1)
          for page in data.get("pages", [])
          for e in page.get("elements", [])
          if e.get("type") == "text" and e.get("text")]
    if not xs:
        return None
    left = min(xs)
    return sum(1 for x in xs if abs(x - left) < 1.0)


def run(title: str, wanted: dict, name) -> int:
    agree, rows = 0, []
    for key, want in sorted(wanted.items()):
        probe = PROBES / name(*key)
        got = oxi_left(probe) if probe.is_file() else None
        rows.append((key, want, got))
        if got == want:
            agree += 1
    print(f"{title}: {agree}/{len(rows)} agree")
    for (a, b), want, got in rows:
        if got != want:
            print(f"  {a:2} lines, {b:2}: word {want}  oxi {got}")
    return len(rows) - agree


def main() -> int:
    # Word gives the same answers at compat 14 as at 15 — measured, all 55 —
    # so the same table serves both. The ENGINE takes a different arm below
    # compat 15, which is the reason to run them separately.
    bad = 0
    for compat in (15, 14):
        bad += run(f"position sweep, compat {compat}", WORD,
                   lambda n, at, c=compat: f"pos_{n}lines_tall24hp_at{at}_compat{c}.docx")
        bad += run(f"size sweep, compat {compat}", SIZES,
                   lambda n, hp, c=compat: f"col_{n}lines_last{hp}hp_compat{c}.docx")
    bad += run("wrapped paragraphs, compat 15", WRAPPED,
               lambda n, chars: f"wrap_{n}paras_{chars}chars_compat15.docx")
    bad += run("closing shape, compat 15", {(k, ""): v for k, v in CLOSING.items()},
               lambda k, _u: f"end_8lines_{k}_compat15.docx")
    bad += run("prefixed run, compat 15", PREFIXED,
               lambda n, pre: f"pre_{n}lines_prefix{pre}_compat15.docx")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
