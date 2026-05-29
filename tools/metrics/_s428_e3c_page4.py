"""S428: investigate e3c545 page-4 top ~11pt offset.

Page 4 of e3c545 is uniformly dy=-12..-14 (Oxi too high); pages 3 and 5
are dy~0. The first element of page 4 (a table row) sits at Word y=68 but
Oxi y=57 (top margin). Find what creates the ~11pt above the page-4 table.
"""
from __future__ import annotations
import os, sys, time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                    "e3c545fac7a7_LOD_Handbook.docx")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(0.3)
    try:
        n = doc.Paragraphs.Count
        print("n_paras", n)
        prev_tbl_id = None
        for pi in range(1, n + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            start = doc.Range(rng.Start, rng.Start)
            page = start.Information(3)
            if page is None or page < 3 or page > 5:
                continue
            y = start.Information(6)
            txt = (rng.Text or "").replace("\r", "").replace("\x07", "↳")[:24]
            try:
                ntab = rng.Tables.Count
            except Exception:
                ntab = 0
            in_tbl = ntab > 0
            # table identity + row/cell
            tinfo = ""
            if in_tbl:
                try:
                    t = rng.Tables(1)
                    # spacing
                    sb = p.SpaceBefore
                    sa = p.SpaceAfter
                    tinfo = f" sb={sb} sa={sa}"
                except Exception as e:
                    tinfo = f" err{e}"
            else:
                sb = p.SpaceBefore
                sa = p.SpaceAfter
                tinfo = f" sb={sb} sa={sa}"
            mark = ""
            print(f"pi={pi:>3} pg{page} y={y:>6.1f} tbl={int(in_tbl)} ntab={ntab}{tinfo}  {txt!r}")
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()


if __name__ == "__main__":
    main()
