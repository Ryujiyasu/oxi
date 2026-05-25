"""COM-measure 29dc6e page 5 paragraphs i=263..i=267 to verify the 11pt
excess gap between Word i=265 and i=266 reported by pagination_word.json.

If COM directly confirms the +11pt anomaly, the bug is real (not a
measurement artifact). Then dump the cell properties of the containing
cell to find what context-dependent factor adds 11pt.

Run from repo root:
  python tools/metrics/com_measure_29dc6e_p5.py
"""
from __future__ import annotations

import os

DOC = "tools/golden-test/documents/docx/29dc6e8943fe_order_01.docx"


def main() -> None:
    import win32com.client as win32

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
        paras = doc.Paragraphs
        print(f"Total paragraphs in 29dc6e: {paras.Count}")
        # Per Word data (pagination_word.json):
        #   i=263 y=376  i=264 y=415  i=265 y=443  i=266 y=482  i=267 y=69.5
        # Verify directly via Word COM, also dump Information(8)=line index.
        print("\nParagraph y measurements (R30 collapsed-start convention):")
        for i in range(260, 275):
            if i < 1 or i > paras.Count:
                continue
            p = paras(i)
            rng = p.Range
            collapsed = doc.Range(rng.Start, rng.Start)
            try:
                y = collapsed.Information(6)  # wdVerticalPositionRelativeToPage
                page = collapsed.Information(1)
                line = collapsed.Information(10)  # wdFirstCharacterLineNumber
            except Exception as e:
                print(f"  para_{i}: error {e}")
                continue
            txt = (rng.Text or "").rstrip("\r")[:50]
            print(f"  para_{i}: page={page} y={y:.2f}pt line_in_page={line}  text={txt!r}")

        # Also: for i=265, measure the END y (last char) to know its content height.
        # Subtract i=265 START from i=265 END = content height.
        # If i=265 takes 22pt, content end is at 443+22=465. Then i=266 starts at 482.
        # Pure spacing between i=265 end and i=266 start = 482-465 = 17pt.
        # Of that, i=266 space-before is 6pt → leftover 11pt mystery.
        print("\ni=265 detailed start/end:")
        p265 = paras(265)
        rng = p265.Range
        y_start = doc.Range(rng.Start, rng.Start).Information(6)
        y_end = doc.Range(rng.End, rng.End).Information(6)
        print(f"  i=265 START y={y_start:.2f}, END y={y_end:.2f}, span={y_end-y_start:.2f}pt")

        p266 = paras(266)
        rng = p266.Range
        y_start = doc.Range(rng.Start, rng.Start).Information(6)
        y_end = doc.Range(rng.End, rng.End).Information(6)
        print(f"  i=266 START y={y_start:.2f}, END y={y_end:.2f}, span={y_end-y_start:.2f}pt")

        # Information for the cell containing i=265.
        rng = p265.Range
        try:
            in_table = rng.Information(12)  # wdWithInTable
            print(f"\ni=265 in table: {in_table}")
            if in_table:
                tbl_y = rng.Cells(1).Range.Information(6)
                print(f"  cell.Range start y={tbl_y:.2f}")
        except Exception as e:
            print(f"  cell inspect error: {e}")

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
