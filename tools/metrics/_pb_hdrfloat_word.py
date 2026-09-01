# -*- coding: utf-8 -*-
"""Word truth + Oxi readback for _pb_hdrfloat_gen.

Prints, per arm, the image's placed top from Word's PDF next to the two
candidate references:

    hdr   = header distance + posOffset   (the header's own first paragraph)
    marg  = top margin      + posOffset   (what Oxi used before S1268b)
"""
import json
import os
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_hdrfloat_gen import ARMS, OUT  # noqa: E402

import fitz  # noqa: E402
import win32com.client  # noqa: E402

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DW = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release",
                  "oxi-dwrite-renderer.exe")
TOP_MARGIN = 1440 / 20.0

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
    print("%-8s %6s %7s %5s | %8s %8s | %8s %8s"
          % ("arm", "hdrTw", "off", "lead", "hdr_y", "marg_y", "word_y", "oxi_y"))
    for tag in ARMS:
        hd, voff, lead = ARMS[tag]
        hdr_y = hd / 20.0 + voff
        marg_y = TOP_MARGIN + voff

        p = os.path.abspath(os.path.join(OUT, tag + ".docx"))
        pdf = p[:-5] + ".pdf"
        d = retry(lambda: word.Documents.Open(p, ReadOnly=True))
        try:
            retry(lambda: d.SaveAs2(pdf, FileFormat=17))
        finally:
            retry(lambda: d.Close(False))
        doc_ = fitz.open(pdf)
        info = doc_[0].get_image_info()
        doc_.close()
        word_y = info[0]["bbox"][1] if info else None

        dj = os.path.join(OUT, tag + "_oxi.json")
        subprocess.run([DW, p, dj[:-5], "110", "--dump-layout=" + dj], capture_output=True)
        oxi_y = None
        if os.path.exists(dj):
            with open(dj, encoding="utf-8") as f:
                lay = json.load(f)
            for e in lay["pages"][0]["elements"]:
                if e["type"] == "image":
                    oxi_y = e["y"]
                    break

        print("%-8s %6d %7.1f %5s | %8.2f %8.2f | %8s %8s"
              % (tag, hd, voff, "yes" if lead else "-", hdr_y, marg_y,
                 "-" if word_y is None else "%.2f" % word_y,
                 "-" if oxi_y is None else "%.2f" % oxi_y))
finally:
    word.Quit()
