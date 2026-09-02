# -*- coding: utf-8 -*-
"""What puts the BLANK page 2 into `reference__0ea3ec86480140c2`?

Word renders p1 (a divider page), a completely empty p2 (0 text, 0 images,
0 drawings), then logical 90 on p3. Oxi renders no blank at all, so 304 of its
1173 paragraphs sit one page early (score 0.7315).

The hand-written `g` arms of `_pb_blankpage_gen.py` -- a continuous section that
restarts numbering -- did NOT reproduce it: Word padded nothing there. So the
cause is something this document has and the minimal one does not, and the way
to find it is to take the REAL file apart one attribute at a time
([[feedback_subtractive_bisection]], [[probe_minimal_docx_degraded]]).

Each arm edits word/document.xml (or settings.xml) of the real file, exports it
through Word, and reports the page count and which pages come out empty.

    python _pb_0ea3_bisect.py
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
SRC = os.path.join(REPO, "pipeline_data", "docx_corpus", "ja", "reference",
                   "0ea3ec86480140c2.docx")
OUT = os.path.join(REPO, "pipeline_data", "_pb_0ea3")

SECT = re.compile(r"<w:sectPr[^>]*>.*?</w:sectPr>", re.S)


def nth_sect_edit(xml, n, fn):
    """Apply `fn` to the n-th (1-based) sectPr's XML."""
    out, i = [], 0
    for k, m in enumerate(SECT.finditer(xml), 1):
        out.append(xml[i:m.start()])
        out.append(fn(m.group(0)) if k == n else m.group(0))
        i = m.end()
    out.append(xml[i:])
    return "".join(out)


def drop_start(t):
    return re.sub(r'\s*w:start="\d+"', "", t)


def drop_type(t):
    return re.sub(r"<w:type w:val=\"\w+\"/>", "", t)


ARMS = {
    "a0_asis": (lambda x: x, lambda s: s),
    "a1_no_sec2_start": (lambda x: nth_sect_edit(x, 2, drop_start), lambda s: s),
    "a2_no_sec3_start": (lambda x: nth_sect_edit(x, 3, drop_start), lambda s: s),
    "a3_no_sec1_start": (lambda x: nth_sect_edit(x, 1, drop_start), lambda s: s),
    "a4_no_sec4_oddpage": (lambda x: nth_sect_edit(x, 4, drop_type), lambda s: s),
    "a5_no_eoh": (lambda x: x, lambda s: s.replace("<w:evenAndOddHeaders/>", "")),
    "a6_no_starts_at_all": (
        lambda x: SECT.sub(lambda m: drop_start(m.group(0)), x), lambda s: s),
}


def build():
    os.makedirs(OUT, exist_ok=True)
    z = zipfile.ZipFile(SRC)
    names = z.namelist()
    xml = z.read("word/document.xml").decode("utf-8")
    st = z.read("word/settings.xml").decode("utf-8")
    for cid, (fx, fs) in ARMS.items():
        dx, ds = fx(xml), fs(st)
        p = os.path.join(OUT, cid + ".docx")
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as o:
            for n in names:
                if n == "word/document.xml":
                    o.writestr(n, dx.encode("utf-8"))
                elif n == "word/settings.xml":
                    o.writestr(n, ds.encode("utf-8"))
                else:
                    o.writestr(n, z.read(n))
        print("built", cid)


def measure():
    import win32com.client as win32
    import fitz
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        print(f"\n{'arm':<22} {'pages':>5}  blanks (0 text, 0 image, 0 drawing)")
        for cid in ARMS:
            src = os.path.join(OUT, cid + ".docx")
            pdf = os.path.join(OUT, cid + ".pdf")
            d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.ExportAsFixedFormat(OutputFileName=pdf, ExportFormat=17)
            finally:
                d.Close(False)
            doc = fitz.open(pdf)
            blanks = [i + 1 for i, pg in enumerate(doc)
                      if not pg.get_text().strip() and not pg.get_images(full=True)
                      and not pg.get_drawings()]
            print(f"{cid:<22} {len(doc):>5}  {blanks[:8]}")
            doc.close()
    finally:
        app.Quit()


if __name__ == "__main__":
    build()
    measure()
