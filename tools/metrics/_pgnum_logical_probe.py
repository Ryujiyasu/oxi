# -*- coding: utf-8 -*-
"""Word's PHYSICAL vs LOGICAL page number, per section, for one document.

S1291 states the blank-page rules in logical page numbers, and Oxi can only
apply them where it can compute that number. A CONTINUOUS section carrying its
own `<w:pgNumType w:start>` is the case it cannot: the S560 merge folds the
section into the previous IR page and the restart, which happens partway down a
page, has nowhere to live. Before carrying it through the merge, this reads what
Word actually does with such a restart.

`wdActiveEndAdjustedPageNumber` (1) is the number Word DISPLAYS -- the logical
one; `wdActiveEndPageNumber` (3) is the physical sheet. Reading both per section
says whether a mid-page restart takes effect on the page it appears on, on the
next one, or not at all.

R30: read them off a COLLAPSED start range.

    python _pgnum_logical_probe.py <docx> [more.docx ...]
"""
import os
import re
import sys
import zipfile

import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

wdActiveEndAdjustedPageNumber = 1
wdActiveEndSectionNumber = 2
wdActiveEndPageNumber = 3


def declared(path):
    """(type, pgNumType start) per sectPr, in document order."""
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "replace")
    out = []
    for m in re.finditer(r"<w:sectPr[^>]*>.*?</w:sectPr>", xml, re.S):
        t = m.group(0)
        ty = re.search(r'<w:type w:val="(\w+)"', t)
        st = re.search(r'<w:pgNumType[^>]*w:start="(\d+)"', t)
        out.append((ty.group(1) if ty else "nextPage",
                    int(st.group(1)) if st else None))
    return out


def retry(fn, tries=40, wait=0.5):
    """Word rejects calls while it is repaginating a long document
    ("Call was rejected by callee"). That is a BUSY signal, not a failure."""
    import time
    last = None
    for _ in range(tries):
        try:
            return fn()
        except Exception as e:  # pywintypes.com_error included
            last = e
            time.sleep(wait)
    raise last


def run(path):
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = app.Documents.Open(os.path.abspath(path), ReadOnly=True, AddToRecentFiles=False)
    try:
        decl = declared(path)
        npages = retry(lambda: int(doc.ComputeStatistics(2)))
        print(f"=== {os.path.basename(path)}  ({npages} pages)")
        print(f"  {'sec':>4} {'declared':>22} {'first para':>10} {'physical':>8} {'logical':>7}")
        seen = set()
        for p in doc.Paragraphs:
            rng = p.Range
            st = doc.Range(rng.Start, rng.Start)
            sec = retry(lambda: int(st.Information(wdActiveEndSectionNumber)))
            if sec in seen:
                continue
            seen.add(sec)
            phys = retry(lambda: int(st.Information(wdActiveEndPageNumber)))
            logi = retry(lambda: int(st.Information(wdActiveEndAdjustedPageNumber)))
            d = decl[sec - 1] if sec - 1 < len(decl) else ("?", None)
            print(f"  {sec:>4} {str(d):>22} {'':>10} {phys:>8} {logi:>7}")
        # every physical page's logical number, read off the first range on it
        print(f"\n  {'physical':>8} {'logical':>7}")
        for pg in range(1, npages + 1):
            try:
                r = retry(lambda: doc.GoTo(What=1, Which=1, Count=pg))
                r = doc.Range(r.Start, r.Start)
                logi = retry(lambda: int(r.Information(wdActiveEndAdjustedPageNumber)))
                sec = retry(lambda: int(r.Information(wdActiveEndSectionNumber)))
                print(f"  {pg:>8} {logi:>7}   sec={sec}")
            except Exception as e:
                print(f"  {pg:>8}   FAIL {str(e)[:50]}")
    finally:
        doc.Close(False)
        app.Quit()


if __name__ == "__main__":
    for a in sys.argv[1:]:
        run(a)
