"""COM-measure ALL d1e8 paragraphs: y_start, y_end, n_lines, font sz."""
import os
import json
import win32com.client as win32

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(REPO_ROOT, "tools", "golden-test", "documents", "docx",
                    "d1e8ac8fd1cc_kyodokenkyuyoushiki06.docx")
OUT = os.path.join(REPO_ROOT, "tools", "metrics", "d1e8_full_word_measure.json")

wdHorizontalPositionRelativeToPage = 5
wdVerticalPositionRelativeToPage = 6
wdActiveEndPageNumber = 3
wdStatisticLines = 1


def first_run_font(p):
    """Return (font_name_east, sz_half_pts) from first run or pPr_rpr."""
    try:
        # First run
        rng = p.Range
        first = rng.Characters(1) if rng.Characters.Count > 0 else None
        if first:
            font = first.Font
            return font.NameFarEast or font.NameAscii, font.Size
    except Exception:
        pass
    return None, None


def main():
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    out = []
    try:
        doc = app.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
        try:
            n = doc.Paragraphs.Count
            print(f"Total paragraphs: {n}")
            for i in range(1, n + 1):
                p = doc.Paragraphs(i)
                rng = p.Range
                r0 = doc.Range(rng.Start, rng.Start)
                y_start = r0.Information(wdVerticalPositionRelativeToPage)
                x_start = r0.Information(wdHorizontalPositionRelativeToPage)
                page = r0.Information(wdActiveEndPageNumber)
                # End
                if rng.End > rng.Start:
                    r1 = doc.Range(rng.End - 1, rng.End - 1)
                else:
                    r1 = doc.Range(rng.Start, rng.Start)
                y_end = r1.Information(wdVerticalPositionRelativeToPage)
                end_page = r1.Information(wdActiveEndPageNumber)
                try:
                    n_lines = rng.ComputeStatistics(wdStatisticLines)
                except Exception:
                    n_lines = None
                fontname, sz = first_run_font(p)
                text = (rng.Text or "").rstrip("\r\x07")[:30]
                entry = {
                    "wi": i,
                    "page": page,
                    "end_page": end_page,
                    "y_start": round(y_start, 2),
                    "y_end": round(y_end, 2),
                    "n_lines": n_lines,
                    "x_start": round(x_start, 2),
                    "font": fontname,
                    "sz_pt": sz,
                    "text": text,
                }
                out.append(entry)
                print(f"  wi={i:>3} p={page}->{end_page} y={y_start:>6.2f}->{y_end:>6.2f} h={y_end-y_start:>5.2f} L={n_lines} sz={sz} fnt={fontname!r} text={text!r}")
        finally:
            doc.Close(SaveChanges=False)
    finally:
        app.Quit()
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
