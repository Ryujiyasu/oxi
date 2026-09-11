# -*- coding: utf-8 -*-
"""Read a two-column page LINE by line, not paragraph by paragraph.

`_col_split_geom.py` asks Word for each paragraph's position, which is enough
while every paragraph is one line. It stops being enough the moment a
paragraph wraps: the split can then fall INSIDE a paragraph, and a paragraph
reports only where it starts.

Word's page object model exposes the layout directly — a page holds
rectangles, a rectangle holds lines, and in a two-column section the text
rectangles ARE the columns. That is what this reads.

    python tools/metrics/_col_split_lines.py <probe-name>

Needs the document open in a window, so it runs Word visible.
"""
import sys
from pathlib import Path

import win32com.client

PROBES = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "column_split"
TEXT_RECTANGLE = 0  # wdTextRectangle — and each one is a LINE, not a column


def main() -> int:
    # --tops also prints every line's top, which is the only way to recover
    # what each line COSTS the column. Rectangle tops are whole points; the
    # sub-point truth needs Information(6), one paragraph at a time.
    tops = "--tops" in sys.argv
    names = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not names:
        print(f"usage: {Path(sys.argv[0]).name} <probe-name> [...]")
        return 2
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    app.DisplayAlerts = False
    try:
        for name in names:
            path = PROBES / f"{name}.docx"
            if not path.is_file():
                print(f"\n{name}: no such probe")
                continue
            doc = app.Documents.Open(str(path), False, False)
            print(f"\n{name}")
            try:
                # Rectangles exist only in print layout; in any other view
                # the collection is silently empty.
                doc.Windows(1).View.Type = 3  # wdPrintView
                pane = doc.Windows(1).Panes(1)
                for pi in range(1, pane.Pages.Count + 1):
                    page = pane.Pages(pi)
                    print(f"  page {pi}: {page.Rectangles.Count} rectangles")
                    seen: dict[int, list] = {}
                    for ri in range(1, page.Rectangles.Count + 1):
                        rect = page.Rectangles(ri)
                        # Word refuses these on rectangles that are not text
                        # (page-number widgets and the like). One that will
                        # not answer is not text.
                        try:
                            if rect.RectangleType != TEXT_RECTANGLE:
                                continue
                            lines = rect.Lines
                            count = lines.Count
                        except Exception:
                            continue
                        # A rectangle is a BLOCK of contiguous lines, not one
                        # line. Counting rectangles undercounts every wrapped
                        # paragraph, which is the whole point of these probes.
                        for li in range(1, count + 1):
                            one = lines(li).Range.Text.replace("\r", "").strip()
                            seen.setdefault(int(rect.Left), []).append(
                                (int(rect.Top), li, int(rect.Width), one))
                    for left in sorted(seen):
                        rows = sorted(seen[left])
                        print(f"    x={left:4}  w={rows[0][2]:4}  {len(rows):3} lines  "
                              f"y {rows[0][0]}..{rows[-1][0]}  "
                              f"{rows[0][3][:18]!r} .. {rows[-1][3][:18]!r}")
                        if tops:
                            print("      tops " + " ".join(str(r[0]) for r in rows))
            finally:
                doc.Close(False)
    finally:
        app.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
