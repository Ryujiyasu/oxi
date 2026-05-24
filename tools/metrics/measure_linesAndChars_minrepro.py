"""Open each docGrid linesAndChars minimal repro in Word and report per-line
chars + rendered span. Companion to ``build_linesAndChars_minrepro.py``.

For each ``tools/metrics/output/linesAndChars_repro/V*.docx`` fixture, walks
every character of paragraph 1 via ``Range.Information(5)`` (horizontal pos)
and ``Information(6)`` (vertical pos) and groups by Y. The output documents
where Word actually wraps and how much horizontal advance each char takes —
needed to distinguish "Oxi compresses chars" from "Oxi has wrong page-break"
hypotheses.

Found 2026-05-24 (S261 investigation): with the doc Normal style applied
(sz=24 = 12pt + jc=both + kern=2), all 6 V0..V5 minimal-repro variants
produce identical Word output: 38 chars / line 1 (12pt fullwidth = 11.6pt
avg). This falsified the "linesAndChars docGrid causes 2× char compression"
hypothesis that the original (stale) Word DML cache had suggested.

Constants reference:
  Information(5) = wdHorizontalPositionRelativeToPage  (points from page left)
  Information(6) = wdVerticalPositionRelativeToPage    (points from page top)

The earlier-tried Information(1) is wdActiveEndAdjustedPageNumber, NOT a
position — that mistake produced the all-same-value output that initially
looked like ``Visible=False`` quirk.

Usage:
  python tools/metrics/build_linesAndChars_minrepro.py   # build fixtures
  python tools/metrics/measure_linesAndChars_minrepro.py  # measure them
"""
import glob
import os
import sys
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WD_HORIZ_POS_REL_PAGE = 5
WD_VERT_POS_REL_PAGE = 6

REPRO_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "output", "linesAndChars_repro")
)

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False

try:
    for path in sorted(glob.glob(os.path.join(REPRO_DIR, "*.docx"))):
        doc = word.Documents.Open(path, ReadOnly=True)
        try:
            para = doc.Paragraphs(1)
            rng = para.Range
            chars = rng.Characters
            n_chars = chars.Count

            lines = []  # (chars_count, x_start, y, last_x)
            cur_count = 0
            cur_x_start = None
            cur_y = None
            cur_last_x = None
            for ci in range(1, n_chars + 1):
                c = chars(ci)
                ch = c.Text
                if ch in ("\r", "\x07", "\x0b"):
                    continue
                cx = c.Information(WD_HORIZ_POS_REL_PAGE)
                cy = c.Information(WD_VERT_POS_REL_PAGE)
                if cur_y is None or abs(cy - cur_y) > 0.5:
                    if cur_count > 0:
                        lines.append((cur_count, cur_x_start, cur_y, cur_last_x))
                    cur_count = 0
                    cur_x_start = cx
                    cur_y = cy
                cur_count += 1
                cur_last_x = cx
            if cur_count > 0:
                lines.append((cur_count, cur_x_start, cur_y, cur_last_x))

            ps = doc.Sections(1).PageSetup
            ta = ps.PageWidth - ps.LeftMargin - ps.RightMargin
            print(
                f"== {os.path.basename(path)} ==  textArea={ta:.1f}pt  total_lines={len(lines)}"
            )
            for i, (cn, xs, y, lx) in enumerate(lines):
                line_w = (lx - xs) if (lx is not None and xs is not None) else 0
                avg = line_w / max(cn - 1, 1) if cn > 1 else 0
                print(
                    f"   L{i+1}: chars={cn:3d}  x_start={xs:.1f}  y={y:.1f}  span={line_w:.1f}pt  avg/ch~{avg:.2f}pt"
                )
        finally:
            doc.Close(SaveChanges=False)
finally:
    word.Quit()
