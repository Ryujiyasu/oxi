"""Measure Word's actual body_bottom_y vs fn area_top gap on b837808 page 1.

R7.29 hypothesis: Oxi's fn-placement code has a +2pt safety buffer
(mod.rs:2552) that rejects fn 4 by 1.4pt on Oxi page 1. If Word's buffer
is smaller (or zero), reducing Oxi's would prevent the spill cascade.

Method: open b837808 via Word COM, walk all paragraphs and footnotes,
find on page 1:
  - body_bottom_y = max(Y + line_height) across body paragraphs
  - fn_top_y = min(Y) across footnotes on page 1
  - gap = fn_top_y - body_bottom_y

Compare with Oxi's measured gap=29.0pt (from OXI_FN_PROBE).
"""
import os
import sys
import win32com.client


def main(docx_rel: str = "tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx"):
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    docx = os.path.normpath(os.path.join(repo, docx_rel))
    if not os.path.exists(docx):
        print(f"docx not found: {docx}", file=sys.stderr)
        return 2

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(docx, ReadOnly=True)
    try:
        # Information constants
        wdActiveEndPageNumber = 3
        wdHorizontalPositionRelativeToPage = 5
        wdVerticalPositionRelativeToPage = 6

        # Body paragraphs on page 1 — also include i=21 (Word puts it on page 1 too)
        print("=== BODY paragraphs on Word page 1 (every paragraph) ===")
        body_bots = []
        for i in range(1, min(doc.Paragraphs.Count, 30) + 1):
            p = doc.Paragraphs(i)
            r = p.Range
            start_r = doc.Range(r.Start, r.Start)
            page = start_r.Information(wdActiveEndPageNumber)
            y_top = start_r.Information(wdVerticalPositionRelativeToPage)
            end_r = doc.Range(r.End - 1, r.End - 1) if r.End > r.Start else r
            try:
                y_end = end_r.Information(wdVerticalPositionRelativeToPage)
                end_page = end_r.Information(wdActiveEndPageNumber)
            except Exception:
                y_end, end_page = y_top, page
            if page <= 2:
                print(f"  para_i={i:3d}  page={page}  y_top={y_top:.2f}  y_end={y_end:.2f}  end_page={end_page}")
                if page == 1 and end_page == 1:
                    body_bots.append((i, y_top, y_end))
        last_body_y_top = body_bots[-1][1] if body_bots else None
        last_body_y_end = body_bots[-1][2] if body_bots else None
        print(f"\n  Last body paragraph END (top of last line) = {last_body_y_end:.2f}")

        # Now check footnotes on page 1
        print("\n=== FOOTNOTES on Word page 1 ===")
        fn_tops = []
        for k in range(1, doc.Footnotes.Count + 1):
            fn = doc.Footnotes(k)
            r = fn.Range
            start_r = doc.Range(r.Start, r.Start)
            page = start_r.Information(wdActiveEndPageNumber)
            if page == 1:
                y = start_r.Information(wdVerticalPositionRelativeToPage)
                ref_r = fn.Reference
                ref_page = ref_r.Information(wdActiveEndPageNumber)
                ref_y = ref_r.Information(wdVerticalPositionRelativeToPage)
                fn_text = fn.Range.Text.strip()[:30]
                print(f"  fn[{k}]  y={y:.2f}  ref_page={ref_page} ref_y={ref_y:.2f}  text='{fn_text}'")
                fn_tops.append((k, y))
        if not fn_tops:
            print("  (no footnotes on page 1)")
        else:
            min_fn_y = min(y for _, y in fn_tops)
            print(f"\n  First (lowest y) fn on page 1: y = {min_fn_y:.2f}")
            print(f"  Last body end y:                  {last_body_y_end:.2f}")
            print(f"  Word's gap (fn_top - body_end):   {min_fn_y - last_body_y_end:.2f} pt")
            print(f"\n  For comparison, Oxi's gap on page 1 = 29.0 pt (OXI_FN_PROBE)")

        # Also look at page 2 for context
        print("\n=== FOOTNOTES on Word page 2 ===")
        for k in range(1, doc.Footnotes.Count + 1):
            fn = doc.Footnotes(k)
            start_r = doc.Range(fn.Range.Start, fn.Range.Start)
            if start_r.Information(wdActiveEndPageNumber) == 2:
                y = start_r.Information(wdVerticalPositionRelativeToPage)
                fn_text = fn.Range.Text.strip()[:30]
                print(f"  fn[{k}]  y={y:.2f}  text='{fn_text}'")

    finally:
        doc.Close(SaveChanges=False)
        word.Quit()
    return 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
