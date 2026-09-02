"""Word COM: per-paragraph (section, page, x, y) for vertical multi-column docs.

Purpose: derive how Word lays out `w:cols num=N` BANDS in a tbRl section, and
what a CONTINUOUS section break with a different N does to the band grid.

R30: Information(2/3/5/6) must be read off a COLLAPSED start range, else the
active-end position is returned and multi-column/multi-page paragraphs report
the wrong band.
"""
import json, os, sys
import win32com.client as win32

wdActiveEndSectionNumber = 2
wdActiveEndPageNumber = 3
wdHorizontalPositionRelativeToPage = 5
wdVerticalPositionRelativeToPage = 6


def probe(path):
    path = os.path.abspath(path)
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = app.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
    try:
        secs = []
        for s in doc.Sections:
            ps = s.PageSetup
            tc = s.PageSetup.TextColumns
            secs.append({
                "index": s.Index,
                "n_cols": tc.Count,
                "spacing": round(float(tc.Spacing), 2),
                "evenly": bool(tc.EvenlySpaced),
                "orient_vert": int(ps.TextColumns.Count) and None,
                "page_w": round(float(ps.PageWidth), 2),
                "page_h": round(float(ps.PageHeight), 2),
                "top": round(float(ps.TopMargin), 2),
                "bottom": round(float(ps.BottomMargin), 2),
                "left": round(float(ps.LeftMargin), 2),
                "right": round(float(ps.RightMargin), 2),
            })
        paras = []
        for i, p in enumerate(doc.Paragraphs, 1):
            rng = p.Range
            start = doc.Range(rng.Start, rng.Start)
            paras.append({
                "i": i,
                "sec": int(start.Information(wdActiveEndSectionNumber)),
                "page": int(start.Information(wdActiveEndPageNumber)),
                "x": round(float(start.Information(wdHorizontalPositionRelativeToPage)), 2),
                "y": round(float(start.Information(wdVerticalPositionRelativeToPage)), 2),
                "n_lines": None,
                "text": rng.Text.rstrip("\r\x07")[:40],
            })
        return {"file": os.path.basename(path), "sections": secs,
                "n_pages": int(doc.ComputeStatistics(2)), "paragraphs": paras}
    finally:
        doc.Close(False)
        app.Quit()


if __name__ == "__main__":
    out = probe(sys.argv[1])
    dest = sys.argv[2] if len(sys.argv) > 2 else None
    if dest:
        with open(dest, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
        print("wrote", dest)
    else:
        print(json.dumps(out, ensure_ascii=False, indent=1)[:4000])
