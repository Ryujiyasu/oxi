"""
Ra2: Body P1 Y offset under various grid pitches and fonts.

Phase 1 found:
  noGrid → delta(P1_y - topMargin) = 0 universally.
  MS Gothic 10.5pt: pitch=360tw(18pt) → delta=+2, pitch=480tw(24pt) → delta=+5.

Hypothesis: delta = (pitch - line_box_inner_height) / 2 (first line centered in grid cell).
This phase pins down line_box_inner for each font/size.

Sweep:
  - linePitch in {280, 320, 360, 400, 440, 480, 560} tw
  - font/size: Calibri 11, MS Gothic 10.5, MS Gothic 14, MS Mincho 10.5, Yu Mincho 10.5
"""
import win32com.client
import json
import os

OUT = os.path.join(os.path.dirname(__file__), "output", "ra2_body_start_grid_sweep.json")
os.makedirs(os.path.dirname(OUT), exist_ok=True)

PITCHES_TW = [280, 320, 360, 400, 440, 480, 520, 560]
FONTS = [
    ("Calibri", 11.0),
    ("Calibri", 14.0),
    ("MS Gothic", 10.5),
    ("MS Gothic", 14.0),
    ("MS Mincho", 10.5),
    ("Yu Mincho", 10.5),
    ("Times New Roman", 11.0),
]

def make_doc(word, font, size, pitch_tw):
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.HeaderDistance = 36
    ps.FooterDistance = 36
    ps.LayoutMode = 2  # wdLayoutModeLineGrid
    body_h = ps.PageHeight - ps.TopMargin - ps.BottomMargin
    ps.LinesPage = max(1, int(round(body_h * 20 / pitch_tw)))

    wdoc.Content.Text = "本文1\r本文2\r本文3"
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = font
        try:
            p.Range.Font.NameFarEast = font
        except Exception:
            pass
        p.Range.Font.Size = size
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0

    wdoc.Repaginate()
    p1y = round(wdoc.Paragraphs(1).Range.Information(6), 4)
    p2y = round(wdoc.Paragraphs(2).Range.Information(6), 4)
    actual_pitch = round((ps.PageHeight - ps.TopMargin - ps.BottomMargin) / ps.LinesPage, 4)
    rec = {
        "font": font, "size": size,
        "set_pitch_tw": pitch_tw, "set_pitch_pt": pitch_tw / 20.0,
        "actual_lines_per_page": ps.LinesPage,
        "actual_pitch_pt": actual_pitch,
        "p1_y": p1y, "p2_y": p2y,
        "delta": round(p1y - 72, 4),
        "p2_minus_p1": round(p2y - p1y, 4),
    }
    wdoc.Close(False)
    return rec

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out = []
    try:
        for font, sz in FONTS:
            print(f"\n--- {font} {sz}pt ---")
            for pt in PITCHES_TW:
                try:
                    r = make_doc(word, font, sz, pt)
                    out.append(r)
                    inner = round(r["actual_pitch_pt"] - 2 * r["delta"], 4)
                    print(f"  pitch={r['set_pitch_pt']:5.2f}pt actual={r['actual_pitch_pt']:6.3f} "
                          f"p2-p1={r['p2_minus_p1']:6.3f} delta={r['delta']:+6.3f} "
                          f"inner_box={inner:6.3f}")
                except Exception as e:
                    print(f"  ERR pitch={pt}: {e}")
    finally:
        word.Quit()
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    print(f"\nSaved {len(out)} to {OUT}")

if __name__ == "__main__":
    main()
