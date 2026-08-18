# -*- coding: utf-8 -*-
"""Probe: does an EMBEDDED font split the line height the same way?

`ascentsplit` settled the rule for installed faces -- the 1.2 line height is
split by the OS/2 win metrics, or by the typo metrics when fsSelection bit 7
(USE_TYPO_METRICS) is set. d28's embedded Calistoga has that bit set and typo
metrics of 1000/-300 on a 1000 em, so the rule predicts a = 0.9231; the single
mixed-size pair available on that deck's title measures 0.937. One pair cannot
settle it and the face is not installed, so this probe asks PowerPoint directly.

Hand-writing a `p:embeddedFontLst` into a python-pptx deck produces a file
PowerPoint refuses to open (the element belongs after `p:notesSz`, which that
template omits). So the deck is built by REUSING d28 through COM: open a copy,
add the arms as new slides, delete the originals. The embedded font parts ride
along untouched, and the file is valid by construction.
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_ROOT = Path(r"pipeline_data\pptx_probes").resolve()
SRC_DIR = Path(r"pipeline_data\pptx_benchmark\dev\pptx").resolve()
DEFAULT_DONOR = "d28"
DEFAULT_FONTS = ["Calistoga", "Jua"]
PAIRS = [(20, 60), (60, 20), (20, 20), (30, 50)]
PP_LAYOUT_BLANK = 12
MSO_TEXT_ORIENT_HORIZONTAL = 1
MSO_TRUE, MSO_FALSE = -1, 0


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--donor", default=DEFAULT_DONOR,
                    help="dev-corpus deck whose embedded fonts to reuse")
    ap.add_argument("--fonts", default=",".join(DEFAULT_FONTS))
    ap.add_argument("--name", default=None, help="probe directory name")
    args = ap.parse_args()
    fonts = [f.strip() for f in args.fonts.split(",")]
    name = args.name or ("embedsplit" if args.donor == DEFAULT_DONOR
                         else f"embedsplit_{args.donor}")
    out = OUT_ROOT / name
    out.mkdir(parents=True, exist_ok=True)
    dst = out / f"{name}.pptx"
    donor = next(SRC_DIR.glob(f"{args.donor}*.pptx"))
    shutil.copyfile(donor, dst)

    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        prs = app.Presentations.Open(str(dst), WithWindow=False)
        try:
            keep = prs.Slides.Count
            for font in fonts:
                for s1, s2 in PAIRS:
                    s = prs.Slides.Add(prs.Slides.Count + 1, PP_LAYOUT_BLANK)
                    cap = s.Shapes.AddTextbox(
                        MSO_TEXT_ORIENT_HORIZONTAL, 18, 9, 500, 24)
                    cap.TextFrame.TextRange.Text = f"{font} {s1}->{s2}"
                    cap.TextFrame.TextRange.Font.Size = 12
                    box = s.Shapes.AddTextbox(
                        MSO_TEXT_ORIENT_HORIZONTAL, 36, 110, 600, 240)
                    tf = box.TextFrame
                    tf.WordWrap = MSO_FALSE
                    tf.AutoSize = 0
                    tf.TextRange.Text = "AAA\rBBB"
                    for idx, pt in ((1, s1), (2, s2)):
                        para = tf.TextRange.Paragraphs(idx)
                        para.Font.Name = font
                        para.Font.Size = pt
                        pf = para.ParagraphFormat
                        pf.LineRuleWithin = MSO_TRUE
                        pf.SpaceWithin = 1.0
                        pf.LineRuleBefore = MSO_FALSE
                        pf.SpaceBefore = 0
                        pf.LineRuleAfter = MSO_FALSE
                        pf.SpaceAfter = 0
            for _ in range(keep):
                prs.Slides(1).Delete()
            prs.Save()
            print(f"wrote {dst}  ({prs.Slides.Count} arms, "
                  f"{args.donor}'s embedded fonts kept)")
        finally:
            prs.Close()
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
