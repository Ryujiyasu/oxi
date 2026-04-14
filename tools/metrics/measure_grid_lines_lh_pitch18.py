"""Measure line height with EXACT pitch=18pt (linePitch=360) grid=lines."""

import win32com.client
import time
import os

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # Use db9c as template (has grid=lines pitch=360)
    # Or create from scratch with explicit settings
    doc = word.Documents.Add()
    time.sleep(0.5)

    sec = doc.Sections(1)
    ps = sec.PageSetup
    ps.TopMargin = 72  # 1 inch
    ps.BottomMargin = 72

    # Set grid: LayoutMode=1 (LineGrid)
    ps.LayoutMode = 1  # wdLayoutModeLineGrid

    # Set exact pitch: 38 lines per page gives 697.9/38=18.37
    # 39 lines: 697.9/39=17.9
    # We need 18.0 exactly. Content height = 841.9 - 72 - 72 = 697.9
    # For pitch=18: 697.9/18 = 38.77 → LinesPage must be 38 (pitch=18.37) or manually set

    # Let me try setting LinesPage to get close to 18
    # 697.9/38 = 18.37, 697.9/39 = 17.9
    # Neither gives exactly 18.
    #
    # Alternative: set margins to make content height divisible by 18
    # content_h = 36*18 = 648. margin = (841.9-648)/2 = 96.95
    # Or: content_h = 38*18 = 684. margin = (841.9-684)/2 = 78.95
    ps.TopMargin = 78.95
    ps.BottomMargin = 78.95
    ps.LinesPage = 38

    content_h = ps.PageHeight - ps.TopMargin - ps.BottomMargin
    actual_pitch = content_h / 38
    print(f"Content height: {content_h:.1f}pt, LinesPage: 38, Pitch: {actual_pitch:.2f}pt")

    # Actually, let me just open db9c which already has the right grid settings
    doc.Close(False)

    doc_path = os.path.abspath("tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=False)
    time.sleep(0.5)

    sec = doc.Sections(1)
    ps = sec.PageSetup
    print(f"\ndb9c settings:")
    print(f"  PageHeight: {ps.PageHeight:.2f}pt")
    print(f"  TopMargin: {ps.TopMargin:.2f}pt")
    print(f"  BottomMargin: {ps.BottomMargin:.2f}pt")
    print(f"  LayoutMode: {ps.LayoutMode}")
    content_h = ps.PageHeight - ps.TopMargin - ps.BottomMargin
    try:
        lp = ps.LinesPage
        print(f"  LinesPage: {lp}")
        print(f"  Actual pitch: {content_h / lp:.2f}pt")
    except:
        print(f"  LinesPage: N/A")
    print(f"  Content height: {content_h:.2f}pt")

    # Measure P8 (4 lines) which is grid-snapped
    p8 = doc.Paragraphs(8)
    r8 = p8.Range
    lines8 = []
    prev_y = None
    for ci in range(r8.Start, min(r8.End, r8.Start + 500)):
        cr = doc.Range(ci, ci + 1)
        cy = cr.Information(6)
        if prev_y is None or abs(cy - prev_y) > 1.0:
            lines8.append(cy)
            prev_y = cy

    print(f"\nP8 line Y positions: {[f'{y:.1f}' for y in lines8]}")
    if len(lines8) > 1:
        gaps = [lines8[i]-lines8[i-1] for i in range(1, len(lines8))]
        print(f"  Gaps: {[f'{g:.1f}' for g in gaps]}")

    # Measure gaps between P1 through P10
    print(f"\nParagraph positions (P1-P15):")
    for i in range(1, 16):
        p = doc.Paragraphs(i)
        rng = p.Range
        start = doc.Range(rng.Start, rng.Start + 1)
        y = start.Information(6)
        chars = len(rng.Text) - 1
        text = rng.Text[:30].replace('\r', '')
        print(f"  P{i}: y={y:.1f} [{chars}c] \"{text}\"")

    doc.Close(False)
    word.Quit()

if __name__ == '__main__':
    measure()
