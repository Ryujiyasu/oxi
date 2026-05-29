"""S429: characterize b837 page 6/7 mid-paragraph break.

Oxi puts 1 line of para i=90 on page 6 + 3 on page 7; page-7 content is a
clean +18.5pt (1 line) too low. Measure where Word actually breaks i=90:
per-line Y via Range char walk, and the page-6 content bottom.
"""
from __future__ import annotations
import os, sys, time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                    "b837808d0555_20240705_resources_data_guideline_02.docx")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(0.3)
    try:
        # page setup
        ps = doc.PageSetup
        ph = ps.PageHeight
        tm = ps.TopMargin
        bm = ps.BottomMargin
        print(f"page_h={ph:.1f} top_margin={tm:.1f} bot_margin={bm:.1f} content_bottom={ph-bm:.1f}")
        # Walk paragraph 90 char-by-char, record visual line Y (Information 6)
        for pi in (89, 90, 91, 92):
            p = doc.Paragraphs(pi)
            rng = p.Range
            txt = (rng.Text or "").replace("\r", "")
            start = rng.Start
            # collect distinct (page, y) per char to find line tops
            seen = []
            n = len(txt) if txt else 0
            step = max(1, n // 40)
            for off in range(0, max(1, n), step):
                r = doc.Range(start + off, start + off)
                pg = r.Information(3)
                y = r.Information(6)
                key = (pg, round(y, 0))
                if not seen or seen[-1] != key:
                    seen.append(key)
            print(f"pi={pi} nchars={n} text={txt[:18]!r}")
            for pg, y in seen:
                print(f"    page {pg}  y={y:.1f}")
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()


if __name__ == "__main__":
    main()
