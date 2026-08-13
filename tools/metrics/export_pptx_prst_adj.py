# -*- coding: utf-8 -*-
"""Export the preset-geometry probe deck to PDF with PowerPoint, and report the
COM view of each shape (type + geometry) so the PDF paths can be attributed.

PowerPoint COM is a singleton: use DispatchEx so we never attach to (or quit) an
instance another session is driving.
"""
import json
import os
import sys

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DIR = os.path.abspath(r"pipeline_data\pptx_probes\prst_adj")
PPTX = os.path.join(DIR, "prst_adj.pptx")
PDF = os.path.join(DIR, "prst_adj.pdf")
TRUTH = os.path.join(DIR, "prst_adj_truth.json")


def main():
    app = win32com.client.DispatchEx("PowerPoint.Application")
    out = []
    try:
        prs = app.Presentations.Open(PPTX, WithWindow=False)
        try:
            for i in range(1, prs.Slides.Count + 1):
                sl = prs.Slides(i)
                shapes = []
                for j in range(1, sl.Shapes.Count + 1):
                    sh = sl.Shapes(j)
                    rec = {
                        "name": sh.Name,
                        "type": int(sh.Type),
                        "left": round(float(sh.Left), 3),
                        "top": round(float(sh.Top), 3),
                        "width": round(float(sh.Width), 3),
                        "height": round(float(sh.Height), 3),
                        "rotation": round(float(sh.Rotation), 3),
                    }
                    try:
                        rec["autoshape"] = int(sh.AutoShapeType)
                    except Exception:
                        rec["autoshape"] = None
                    try:
                        rec["flip_h"] = bool(sh.HorizontalFlip)
                    except Exception:
                        pass
                    shapes.append(rec)
                out.append({"slide": i, "shapes": shapes})
            prs.SaveAs(PDF, 32)
            prs.Close()
        finally:
            app.Quit()
    except Exception:
        try:
            app.Quit()
        except Exception:
            pass
        raise
    with open(TRUTH, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=1, ensure_ascii=False)
    print("wrote %s" % PDF)
    for s in out:
        auto = [x["autoshape"] for x in s["shapes"]]
        print("  slide %2d  shapes=%d  autoshape=%s" % (s["slide"], len(s["shapes"]), auto))


if __name__ == "__main__":
    main()
