"""Measure text Y offset for various font sizes on docGrid lines pitch=360.

Create fresh document with lines grid, write single paragraphs at various sizes,
measure Y position to determine text_y_offset = text_y - grid_line.
"""
import win32com.client
import time

def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Add()
    time.sleep(1)

    # Set docGrid lines mode
    doc.PageSetup.LayoutMode = 1  # wdLayoutModeLineGrid
    doc.PageSetup.TopMargin = 72.0  # Clean 72pt = 4*18 grid-aligned

    sel = word.Selection

    # Test various font/size combinations
    tests = [
        ("ＭＳ ゴシック", 8),
        ("ＭＳ ゴシック", 9),
        ("ＭＳ ゴシック", 10),
        ("ＭＳ ゴシック", 10.5),
        ("ＭＳ ゴシック", 11),
        ("ＭＳ ゴシック", 12),
        ("ＭＳ ゴシック", 14),
        ("Century", 10.5),
        ("Century", 12),
        ("Century", 14),
        ("ＭＳ 明朝", 10.5),
        ("ＭＳ 明朝", 12),
        ("Calibri", 10.5),
        ("Calibri", 11),
    ]

    for fn, sz in tests:
        sel.Font.Name = fn
        sel.Font.Size = sz
        sel.TypeText(f"Test {fn} {sz}pt")
        sel.TypeParagraph()

    time.sleep(1)

    print(f"LayoutMode: {doc.PageSetup.LayoutMode}")
    print(f"TopMargin: {doc.PageSetup.TopMargin}")
    print()

    paras = doc.Paragraphs
    print(f"{'P':>3s} {'Y':>8s} {'gap':>6s} {'grid_n':>7s} {'offset':>7s} {'font':>15s} {'sz':>5s}")

    base = doc.PageSetup.TopMargin
    pitch = 18.0
    prev_y = None

    for i in range(1, min(paras.Count + 1, len(tests) + 1)):
        p = paras(i)
        r = p.Range
        y = r.Information(6)
        fn = r.Font.Name
        sz = r.Font.Size
        gap = y - prev_y if prev_y else 0

        # Grid line (floor)
        grid_n = (y - base) / pitch
        grid_floor = base + int((y - base) / pitch) * pitch
        offset = y - grid_floor

        print(f"P{i:2d} {y:8.2f} {gap:6.2f} {grid_n:7.3f} {offset:7.2f} {fn:>15s} {sz:5.1f}")
        prev_y = y

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
