"""Word COM measurement for the S299 Stage 3 afterLines hypothesis.

Opens both repros from `tools/fixtures/phase2_afterlines_samples/` in
Word, measures the Y of paragraph 1 START and paragraph 2 START, and
prints the gap. Compare across the two repros:

  control_no_afterlines.docx   → expected gap = 28pt (2*11 + 6)
  p2_afterlines_after.docx     → if gap > 28pt: Word treats `w:afterLines+w:after` of
                                   paragraph N+1 as advance-reservation BEFORE that paragraph.
                                   The excess (~11pt expected) is the S299 Stage 3 bug.

Oxi has already been measured (S299): both repros produce gap=28pt in Oxi.
→ If Word's gap differs, the divergence is purely Oxi's missing handling
  of advance-attributed `afterLines+after`. That's a surgical fix domain.

Run from repo root:
  python tools/metrics/com_measure_afterlines_repro.py

Requires Word + pywin32 (`pip install pywin32`).
"""
from __future__ import annotations

import os
import sys


FIXTURE_DIR = os.path.join(
    os.path.dirname(__file__), "..", "fixtures", "phase2_afterlines_samples"
)


def measure(doc_path: str) -> dict:
    import win32com.client as win32

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
        paras = doc.Paragraphs
        # Word reports paragraph y via Range.Information(6) =
        # wdVerticalPositionRelativeToPage. R30 (CLAUDE.md): for paragraphs
        # spanning lines, must collapse the range to its start first.
        result = {}
        for i in range(1, paras.Count + 1):
            p = paras(i)
            rng = p.Range
            # Skip empty paragraphs and table cell-end markers (heuristic).
            txt = (rng.Text or "").strip()
            if not txt:
                continue
            collapsed = doc.Range(rng.Start, rng.Start)
            y = collapsed.Information(6)  # wdVerticalPositionRelativeToPage, pt
            page = collapsed.Information(1)  # wdActiveEndPageNumber
            result[i] = {"y": float(y) / 1.0, "page": int(page), "text": txt[:40]}
        doc.Close(SaveChanges=False)
        return result
    finally:
        word.Quit()


def main() -> None:
    print(f"Measuring afterLines repros from {FIXTURE_DIR}/\n")
    for name in ["control_no_afterlines.docx", "p2_afterlines_after.docx"]:
        path = os.path.join(FIXTURE_DIR, name)
        if not os.path.exists(path):
            print(f"  MISSING: {path}")
            continue
        print(f"=== {name} ===")
        rs = measure(path)
        ys = sorted(rs.items(), key=lambda kv: (kv[1]["page"], kv[1]["y"]))
        prev_y = None
        for i, info in ys:
            gap_str = ""
            if prev_y is not None:
                gap = info["y"] - prev_y
                gap_str = f"  gap_from_prev={gap:+.2f}pt"
            print(f'  para_{i}: page={info["page"]} y={info["y"]:.2f}pt'
                  f'  text={info["text"]!r}{gap_str}')
            prev_y = info["y"]
        print()

    # Expected interpretation:
    #   If control gap_p1->p2 == 28pt AND test gap_p1->p2 > 28pt
    #     → S299 Stage 3 hypothesis CONFIRMED. Excess gap = Word's
    #       advance-attribution of P2's afterLines+after. Bug repro'd
    #       in isolation. Implement fix in Oxi's layout space-before
    #       handling (e.g., add para_n.style.{after_lines,after} to
    #       cursor advance BEFORE laying out para_n+1 when those are
    #       set on para_n+1 itself).
    #   If both repros give gap == 28pt
    #     → S299 Stage 3 hypothesis FALSIFIED. The 29dc6e 11pt excess
    #       has another cause. Re-derive from richer input space.


if __name__ == "__main__":
    main()
