# -*- coding: utf-8 -*-
"""Produce a Word truth PDF for one docx (SaveAs2 FileFormat=17).

Usage: _pdf_truth.py <docx> [out.pdf]
Some corpus docs have a COM pagination JSON but no PDF, and the JSON carries
y=null -- so any question about WHERE on the page something sits needs this.
"""
import os, shutil, sys, tempfile, time
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
docx = os.path.abspath(sys.argv[1])
pdf = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else docx[:-5] + ".pdf"
import win32com.client
tmp = None
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False; word.DisplayAlerts = 0
def retry(fn, tries=8):
    for i in range(tries):
        try: return fn()
        except Exception:
            if i == tries - 1: raise
            time.sleep(2)
try:
    # Tracked changes must not reach the export. Word's markup mode is a sticky
    # APPLICATION preference, so an unpinned export is not reproducible:
    # 0010437a7f75f636 exports 628 pages with All Markup and 625 with none, and
    # no markup is what the renderers draw (S483 sets ShowRevisions::Final).
    # Accept the revisions rather than hide them through the view -- hiding them
    # leaves Information(3) on the markup-shown pagination (see
    # measure_pagination_word.open_clean).
    #
    # Accepting needs the document open read-write, so it is done on a COPY: the
    # corpus file is a source of truth and must not be exposed to a stray save,
    # a lock file or an autosave. A revision-free document keeps the read-only
    # path on the original, unchanged.
    d = retry(lambda: word.Documents.Open(docx, ReadOnly=True))
    try:
        nrev = d.Revisions.Count
    except Exception:
        nrev = 0
    if nrev:
        retry(lambda: d.Close(False))
        fd, tmp = tempfile.mkstemp(suffix=".docx", prefix="pdftruth_")
        os.close(fd)
        shutil.copy(docx, tmp)
        d = retry(lambda: word.Documents.Open(tmp, ReadOnly=False))
        try:
            d.Revisions.AcceptAll()
            d.Repaginate()
            time.sleep(1.0)
        except Exception:
            pass
    try:
        retry(lambda: d.SaveAs2(pdf, FileFormat=17))
    finally:
        retry(lambda: d.Close(False))
finally:
    word.Quit()
    if tmp:
        try:
            os.remove(tmp)
        except Exception:
            pass
print("wrote", pdf, os.path.getsize(pdf), "bytes")
