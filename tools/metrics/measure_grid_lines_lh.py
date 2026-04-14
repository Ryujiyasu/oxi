"""Measure line height for various fonts/sizes in grid=lines mode.
Creates a minimal document with grid=lines, pitch=18, and measures Y gaps."""

import win32com.client
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Add()
    time.sleep(0.5)

    # Set up: A4, grid=lines, linePitch=360 (18pt)
    sec = doc.Sections(1)
    ps = sec.PageSetup
    ps.TopMargin = 72  # 1 inch = 72pt
    ps.BottomMargin = 72

    # Enable grid: lines mode, pitch=18pt (360 twips)
    doc.Sections(1).PageSetup.LayoutMode = 1  # wdLayoutModeLineGrid

    # Set line pitch
    # LayoutMode=1 enables grid. LinePitch in twips.
    # Use PageSetup.LineBetween doesn't exist... use different approach
    # Actually set via XML:
    # doc.Sections(1).PageSetup... doesn't expose linePitch directly.
    # Instead, use content controls or Word object model.

    # Alternative: use doc.PageSetup to set grid
    try:
        ps.LinesPage = 40  # This sets the number of lines per page
        # linePitch = (pageHeight - topMargin - bottomMargin) / linesPage
        content_h = ps.PageHeight - ps.TopMargin - ps.BottomMargin
        actual_pitch = content_h / 40
        print(f"Page content height: {content_h:.1f}pt")
        print(f"Lines per page: 40")
        print(f"Actual pitch: {actual_pitch:.2f}pt")
    except Exception as e:
        print(f"LinesPage failed: {e}")

    # Test fonts/sizes
    test_cases = [
        ("Times New Roman", 12),
        ("Times New Roman", 10.5),
        ("Times New Roman", 11),
        ("Calibri", 11),
        ("Calibri", 10.5),
        ("MS Gothic", 10.5),
    ]

    for font_name, font_size in test_cases:
        # Clear document
        doc.Content.Delete()
        time.sleep(0.3)

        # Add 5 identical paragraphs
        for i in range(5):
            rng = doc.Content
            rng.Collapse(0)  # wdCollapseEnd
            if i > 0:
                rng.InsertParagraphAfter()
                rng = doc.Content
                rng.Collapse(0)
            rng.InsertAfter(f"Test line {i+1} with {font_name} {font_size}pt")

        # Set font for all content
        doc.Content.Font.Name = font_name
        doc.Content.Font.Size = font_size

        time.sleep(0.3)

        # Measure Y positions
        ys = []
        for i in range(1, min(doc.Paragraphs.Count + 1, 6)):
            p = doc.Paragraphs(i)
            rng = p.Range
            start = doc.Range(rng.Start, rng.Start + 1)
            y = start.Information(6)
            ys.append(y)

        gaps = [ys[i] - ys[i-1] for i in range(1, len(ys))]
        avg_gap = sum(gaps) / len(gaps) if gaps else 0

        print(f"\n{font_name} {font_size}pt:")
        print(f"  Y positions: {[f'{y:.1f}' for y in ys]}")
        print(f"  Gaps: {[f'{g:.1f}' for g in gaps]}")
        print(f"  Average gap: {avg_gap:.2f}pt")

    doc.Close(False)
    word.Quit()

if __name__ == '__main__':
    measure()
