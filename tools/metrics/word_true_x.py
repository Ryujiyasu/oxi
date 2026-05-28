"""S417: calibrated GetPoint-based TRUE rendered-x measurement for Word.

S416 proved Range.Information(5) returns a LEFT-FLOW LOGICAL x, not the
rendered glyph position — wrong for any non-left-aligned text. This module
gets the TRUE rendered x via ActiveWindow.GetPoint (screen px) with a
per-page px->pt calibration.

Calibration: for LEFT (align=0) or JUSTIFY (align=3) paragraphs the first
glyph sits at the line's logical left edge, so for those the rendered char0
== Information(5). We collect such (GetPoint_px, Information5_pt) pairs per
page and linear-fit px->pt. (Validated: slope == 0.75 pt/px at 100% zoom /
96 DPI.) Right/center cells are then converted from GetPoint px to pt.

Validated on 1ec1 col4 "（4×U+3000）税": true x = 346.5pt (matches Oxi ON
S412 gate 346.55; Information(5) said 315.75 = artifact).

Usage:
    from word_true_x import measure_true_x
    recs = measure_true_x(docx_path)   # list of {page,y,x_true,x_info5,align,text,in_table}

Multipage: uses ScrollIntoView before each GetPoint. Per-page calibration
re-derived from that page's anchors; falls back to global slope 0.75 + 1
anchor if a page has <2 anchors.
"""
from __future__ import annotations
import sys

GLOBAL_SLOPE = 0.75  # pt/px at 100% zoom, 96 DPI (validated S416)

wdHorizontalPositionRelativeToPage = 5
wdVerticalPositionRelativeToPage = 6
wdActiveEndPageNumber = 3


def _fit(anchors):
    """anchors: list of (px, pt). Return (slope, intercept_px0_pt)."""
    if len(anchors) >= 2:
        (px1, pt1) = anchors[0]
        (px2, pt2) = anchors[-1]
        if px2 != px1:
            slope = (pt2 - pt1) / (px2 - px1)
            return slope, pt1 - slope * px1
    if len(anchors) == 1:
        px1, pt1 = anchors[0]
        return GLOBAL_SLOPE, pt1 - GLOBAL_SLOPE * px1
    return None


def measure_true_x(docx_path: str, max_paras: int | None = None) -> list[dict]:
    import win32com.client as wc
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = True  # GetPoint requires a visible window
    word.ScreenUpdating = False
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    doc.Repaginate()
    win = doc.ActiveWindow
    recs = []
    try:
        n = doc.Paragraphs.Count
        if max_paras:
            n = min(n, max_paras)
        # Pass 1: gather raw (page, align, info5_x, y, px, text, in_table)
        raw = []
        for pi in range(1, n + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            txt = (rng.Text or '').rstrip('\r\n\x07')
            if not txt:
                continue
            s = rng.Start
            start_rng = doc.Range(s, s)
            try:
                page = int(start_rng.Information(wdActiveEndPageNumber))
                y = round(float(start_rng.Information(wdVerticalPositionRelativeToPage)), 2)
                info5 = round(float(start_rng.Information(wdHorizontalPositionRelativeToPage)), 2)
            except Exception:
                continue
            align = p.Format.Alignment  # 0 left 1 center 2 right 3 justify
            in_table = rng.Tables.Count > 0
            # GetPoint (true rendered) for the first char
            try:
                r0 = doc.Range(s, s + 1)
                win.ScrollIntoView(r0, True)
                px = win.GetPoint(0, 0, 0, 0, r0)[0]
            except Exception:
                px = None
            raw.append({'page': page, 'align': align, 'x_info5': info5,
                        'y': y, 'px': px, 'text': txt[:40], 'in_table': in_table})

        # Pass 2: per-page calibration from align in {0,3} (rendered==info5)
        from collections import defaultdict
        anchors_by_page = defaultdict(list)
        for r in raw:
            if r['px'] is not None and r['align'] in (0, 3):
                anchors_by_page[r['page']].append((r['px'], r['x_info5']))
        fits = {}
        for pg, anch in anchors_by_page.items():
            anch_sorted = sorted(anch)
            f = _fit(anch_sorted)
            if f:
                fits[pg] = f

        for r in raw:
            x_true = None
            if r['px'] is not None:
                f = fits.get(r['page'])
                if f is None and r['align'] in (0, 3):
                    # this para is itself a left/justify; rendered == info5
                    x_true = r['x_info5']
                elif f is not None:
                    slope, inter = f
                    x_true = round(slope * r['px'] + inter, 2)
            recs.append({**r, 'x_true': x_true})
    finally:
        doc.Close(SaveChanges=0)
        word.Quit()
    return recs


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    import os
    path = sys.argv[1] if len(sys.argv) > 1 else \
        r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\1ec1091177b1_006.docx'
    recs = measure_true_x(path)
    # Validate: print right/center cells with x_true vs x_info5
    print(f'{os.path.basename(path)}: {len(recs)} paragraphs')
    print(f"{'pg':3s} {'align':5s} {'x_true':8s} {'x_info5':8s} {'Δ':7s} text")
    for r in recs:
        if r['align'] in (1, 2) and r['x_true'] is not None:
            d = r['x_true'] - r['x_info5']
            print(f"{r['page']:<3} {r['align']:<5} {r['x_true']:<8} {r['x_info5']:<8} {d:<7.2f} {r['text'][:24]!r}")
