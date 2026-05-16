"""COM-measure d1e8's wi=37+ paragraph line counts and Y positions."""
import os
import win32com.client as win32

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(REPO_ROOT, "tools", "golden-test", "documents", "docx",
                    "d1e8ac8fd1cc_kyodokenkyuyoushiki06.docx")

wdVerticalPositionRelativeToPage = 6
wdHorizontalPositionRelativeToPage = 5
wdStatisticLines = 1


def main():
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        doc = app.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
        try:
            print(f"Total paragraphs: {doc.Paragraphs.Count}\n")
            # Focus on wi=28..48 (page 2 region)
            for i in list(range(28, 49)):
                if i > doc.Paragraphs.Count:
                    break
                p = doc.Paragraphs(i)
                rng = p.Range
                r0 = doc.Range(rng.Start, rng.Start)
                y_start = r0.Information(wdVerticalPositionRelativeToPage)
                page = r0.Information(3)
                # Range.ComputeStatistics for line count
                try:
                    n_lines = rng.ComputeStatistics(wdStatisticLines)
                except Exception as e:
                    n_lines = "?"
                # text (first 30 chars)
                text = (rng.Text or "").rstrip("\r\x07")[:25]
                # end y
                r1 = doc.Range(max(rng.End - 1, rng.Start), max(rng.End - 1, rng.Start))
                y_end = r1.Information(wdVerticalPositionRelativeToPage)
                end_page = r1.Information(3)
                print(f"  wi={i:>3} page={page}->{end_page} y_start={y_start:>7.2f} y_end={y_end:>7.2f} n_lines={n_lines} text='{text}'")
        finally:
            doc.Close(SaveChanges=False)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
