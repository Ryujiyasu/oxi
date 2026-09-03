"""Export a deck's truth PDF from a WARM PowerPoint.

The corpus's truth PDFs are all COLD -- each was exported from the first open
of its deck (`pptx_truth_pdf_first_open_is_cold`). For a deck whose fonts the
machine resolves differently on the second open, the cold PDF is not the
rendering a reader sees. Deck 47 is such a deck: its cold PDF embeds
`Caladea-Regular` with `ItalicAngle -9` and the ITALIC file's advances
(548/540/507 for Y/C/o), while live PowerPoint measured through COM agrees
with the UPRIGHT file (570/562/531).

Opens the deck once and closes it, so whatever the first open installs is
installed, then reopens and exports. Prints what the resulting PDF says about
each Caladea face so the two states can be compared without opening the PDF by
hand.

    python tools/metrics/pptx_warm_export.py 47 [--out <path>]

Never run this while a pptx render is in flight (`pptx_com_render_must_not_overlap`).
"""
import argparse
import glob
import os
import re
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SSIM = os.path.join(ROOT, "pipeline_data", "pptx_benchmark", "ssim_pptx")


def deck_path(doc: str) -> str:
    hits = glob.glob(os.path.join(ROOT, "pipeline_data", "pptx_benchmark", "pptx", doc + "*.pptx"))
    if not hits:
        sys.exit("no deck matching " + doc)
    return hits[0]


def describe(pdf_path: str) -> None:
    import pymupdf

    pdf = pymupdf.open(pdf_path)
    seen = set()
    for pno in range(len(pdf)):
        for f in pdf[pno].get_fonts(full=True):
            xref, name = f[0], f[3]
            if xref in seen:
                continue
            seen.add(xref)
            d = pdf.xref_object(xref)
            fc = re.search(r"/FirstChar (\d+)", d)
            if not fc:
                continue
            first = int(fc.group(1))
            m = re.search(r"/Widths (\d+) 0 R", d)
            arr = pdf.xref_object(int(m.group(1))) if m else (
                re.search(r"/Widths\s*(\[.*?\])", d, re.S) or re.match("", "")).group(1)
            nums = [int(float(x)) for x in re.findall(r"[-\d.]+", arr)]
            fd = re.search(r"/FontDescriptor (\d+) 0 R", d)
            desc = pdf.xref_object(int(fd.group(1))) if fd else ""
            ia = re.search(r"/ItalicAngle ([-\d.]+)", desc)

            def wid(ch: str):
                i = ord(ch) - first
                return nums[i] if 0 <= i < len(nums) else None

            print("   %-34s ItalicAngle=%-5s Y=%s C=%s o=%s"
                  % (name, ia.group(1) if ia else "?", wid("Y"), wid("C"), wid("o")))


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("doc")
    ap.add_argument("--out", default=None)
    ap.add_argument("--cold", action="store_true", help="skip the warming open")
    args = ap.parse_args()

    import win32com.client

    src = deck_path(args.doc)
    out = args.out or os.path.join(SSIM, "ppt_pdf", args.doc + "_warm.pdf")
    print("deck:", os.path.basename(src))

    app = win32com.client.Dispatch("PowerPoint.Application")
    if not args.cold:
        pres = app.Presentations.Open(src, WithWindow=False)
        time.sleep(3)
        pres.Close()
        time.sleep(1)
        print("warmed (opened and closed once)")

    pres = app.Presentations.Open(src, WithWindow=False)
    time.sleep(2)
    if os.path.exists(out):
        os.remove(out)
    pres.SaveAs(out, 32)  # ppSaveAsPDF
    pres.Close()
    app.Quit()
    print("wrote", out)
    describe(out)
    old = os.path.join(SSIM, "ppt_pdf", args.doc + ".pdf")
    if os.path.exists(old):
        print("the corpus's existing (cold) truth says:")
        describe(old)


if __name__ == "__main__":
    main()
