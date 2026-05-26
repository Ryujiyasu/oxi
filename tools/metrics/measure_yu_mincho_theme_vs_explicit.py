"""S323 COM measurement for Yu Mincho theme vs explicit repro.

For each fixture, opens it in Word, measures Y of paragraphs 1, 2, 3
via `Range.Information(6)` (wdVerticalPositionRelativeToPage), and
prints stride between consecutive paragraphs.

Hypothesis check:
  - All MS Mincho (no Yu involvement): strides should all be ~14.5pt
    or whatever Word uses for sz=11pt MS Mincho.
  - Theme-resolved Yu Mincho middle paragraph: if hypothesis HOLDS,
    stride p1→p2 = stride p2→p3 = ~14.5pt (no 83/64 applied).
  - Explicit Yu Mincho middle paragraph: if hypothesis HOLDS,
    stride p1→p2 = ~14.5pt but stride p2→p3 = ~18.5pt (Yu Mincho
    line height taller).

Reports the actual values so the hypothesis can be falsified or
confirmed without further speculation.
"""
import os
import sys
import win32com.client

FIXTURES = [
    "v1_ms_mincho_only.docx",
    "v1_yu_mincho_theme.docx",
    "v1_yu_mincho_explicit.docx",
]

FIXTURE_DIR = os.path.join(
    os.path.dirname(__file__), "..", "fixtures",
    "yu_mincho_theme_vs_explicit",
)

WD_VERT_POS_REL_TO_PAGE = 6  # wdVerticalPositionRelativeToPage


def measure(path: str) -> list:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
        # Use collapsed-start range per CLAUDE.md R30 fix to avoid
        # active-end-position bug for multi-page paragraphs.
        ys = []
        for i in range(1, doc.Paragraphs.Count + 1):
            rng = doc.Paragraphs(i).Range
            collapsed = doc.Range(rng.Start, rng.Start)
            y = collapsed.Information(WD_VERT_POS_REL_TO_PAGE)
            text = (rng.Text or "").rstrip("\r\n").strip()[:30]
            ys.append((i, y, text))
        doc.Close(False)
        return ys
    finally:
        word.Quit()


def main():
    print(f"{'='*60}")
    print(f"S323 Yu Mincho theme vs explicit COM measurement")
    print(f"{'='*60}")
    for fname in FIXTURES:
        path = os.path.join(FIXTURE_DIR, fname)
        if not os.path.exists(path):
            print(f"SKIP {fname}: not found at {path}")
            continue
        print(f"\n--- {fname} ---")
        ys = measure(path)
        for i, y, text in ys:
            print(f"  para {i}: y={y:6.2f}pt  text={text!r}")
        # Compute strides
        if len(ys) >= 2:
            print(f"  strides:")
            for j in range(1, len(ys)):
                stride = ys[j][1] - ys[j-1][1]
                print(f"    p{ys[j-1][0]} -> p{ys[j][0]}: {stride:6.2f}pt")


if __name__ == "__main__":
    main()
