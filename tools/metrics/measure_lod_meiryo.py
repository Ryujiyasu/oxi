"""Measure Meiryo line height in LOD_Handbook.docx.

Oxi: 20.0pt per line, Word: 20.5pt per line.
Need to confirm exact line height for Meiryo 12pt.
"""
import win32com.client
import os, time

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx", "e3c545fac7a7_LOD_Handbook.docx"))

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    ps = doc.Paragraphs
    print(f"Total paragraphs: {ps.Count}")

    # Measure first 20 paragraphs
    for i in range(1, min(21, ps.Count + 1)):
        p = ps(i)
        r = p.Range
        y = r.Information(6)  # wdVerticalPositionRelativeToPage
        page = r.Information(3)
        fmt = p.Format
        ls = fmt.LineSpacing
        ls_rule = fmt.LineSpacingRule
        sb = fmt.SpaceBefore
        sa = fmt.SpaceAfter
        sz = r.Font.Size
        fn = r.Font.Name
        text = r.Text[:30].replace('\r', '').replace('\x07', '')
        print(f"P{i:3d}: page={page} y={y:7.2f} sz={sz} font={fn} ls={ls:.2f} rule={ls_rule} sb={sb:.1f} sa={sa:.1f} text={text}")

    # Also measure a fresh Meiryo paragraph
    print("\n--- Meiryo metrics check ---")
    # Get GDI line height for Meiryo 12pt
    p1 = ps(1)
    r1 = p1.Range
    print(f"P1 font: {r1.Font.Name} {r1.Font.Size}pt")
    print(f"P1 lineSpacing: {p1.Format.LineSpacing}")
    print(f"P1 lineSpacingRule: {p1.Format.LineSpacingRule}")

    # Get Y difference between P3 and P4 (both are body text)
    p3 = ps(3)
    p4 = ps(4)
    y3 = p3.Range.Information(6)
    y4 = p4.Range.Information(6)
    print(f"\nP3 y={y3:.2f}, P4 y={y4:.2f}, gap={y4-y3:.2f}")

    p5 = ps(5)
    y5 = p5.Range.Information(6)
    print(f"P5 y={y5:.2f}, gap from P4={y5-y4:.2f}")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
