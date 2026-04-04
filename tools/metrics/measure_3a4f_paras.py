"""Measure paragraph Y positions for 3a4f, comparing with Oxi layout."""
import win32com.client
import os, time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    path = os.path.abspath("tools/golden-test/documents/docx/3a4f9fbe1a83_001620506.docx")
    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(1)

    # Measure first 50 paragraphs
    print(f"Total paragraphs: {doc.Paragraphs.Count}")
    print(f"{'P#':>4s} {'page':>5s} {'y':>8s} {'fs':>5s} {'gap':>7s} text")
    print("-" * 80)

    prev_y = None
    prev_page = None
    for i in range(1, min(51, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(i)
        rng = p.Range
        y = rng.Information(6)  # wdVerticalPositionRelativeToPage
        page = rng.Information(3)  # wdActiveEndPageNumber
        fs = p.Range.Font.Size
        text = rng.Text[:30].replace('\r', '').replace('\n', '')

        gap = ""
        if prev_y is not None and page == prev_page:
            gap = f"{y - prev_y:+.1f}"
        elif prev_page is not None and page != prev_page:
            gap = "NEW_PAGE"

        print(f"{i:4d} {page:5d} {y:8.2f} {fs:5.1f} {gap:>7s} {text}")
        prev_y = y
        prev_page = page

    doc.Close(SaveChanges=False)
    word.Quit()

if __name__ == "__main__":
    measure()
