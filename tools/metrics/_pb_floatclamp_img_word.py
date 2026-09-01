# -*- coding: utf-8 -*-
"""Word truth for _pb_floatclamp_img_gen: where does Word place the image?

pymupdf's get_image_info reports the placed rect straight from the page's
content stream, so the drawn origin is read without a rasterisation step.

A first attempt rasterised the page and took the bounding box of the dark
pixels.  It reported the SAME origin for all six arms, because the body text
is darker and further left than the image's black quadrant -- the arms all
agreed on a number that was not the image at all.  Measure the placement, not
the ink, when the page carries other ink.
"""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_floatclamp_img_gen import ARMS, OUT  # noqa: E402

import fitz  # noqa: E402
import win32com.client  # noqa: E402

PW, PH = 11906 / 20.0, 16838 / 20.0
COL_LEFT = 1440 / 20.0

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0


def retry(fn, tries=10):
    for i in range(tries):
        try:
            return fn()
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(1.5)


try:
    print("%-10s %7s %7s | %8s %8s | %8s %8s | %8s %8s"
          % ("arm", "imgW", "imgH", "raw_x", "raw_y", "clmp_x", "clmp_y", "img_x", "img_y"))
    for tag in ARMS:
        w, h, hrel, hoff, vrel, voff = ARMS[tag]
        ref_left = 0.0 if hrel == "page" else COL_LEFT
        raw_x, raw_y = ref_left + hoff, voff
        clmp_x = max(0.0, PW - w) if raw_x + w > PW else raw_x
        clmp_y = max(0.0, PH - h) if raw_y + h > PH else raw_y

        p = os.path.join(OUT, tag + ".docx")
        pdf = p[:-5] + ".pdf"
        d = retry(lambda: word.Documents.Open(os.path.abspath(p), ReadOnly=True))
        try:
            retry(lambda: d.SaveAs2(os.path.abspath(pdf), FileFormat=17))
        finally:
            retry(lambda: d.Close(False))

        doc_ = fitz.open(pdf)
        boxes = doc_[0].get_image_info()
        doc_.close()
        if boxes:
            bb = boxes[0]["bbox"]
            print("%-10s %7.1f %7.1f | %8.2f %8.2f | %8.2f %8.2f | %8.2f %8.2f"
                  % (tag, w, h, raw_x, raw_y, clmp_x, clmp_y, bb[0], bb[1]))
        else:
            print("%-10s %7.1f %7.1f | %8.2f %8.2f | %8.2f %8.2f | %8s %8s"
                  % (tag, w, h, raw_x, raw_y, clmp_x, clmp_y, "-", "-"))
finally:
    word.Quit()
