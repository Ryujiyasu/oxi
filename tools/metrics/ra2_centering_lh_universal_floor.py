"""§3.3 / §13.4 grid centering line height — universal 83/64 floor verification.

Memory `spec_grid_centering_lh_2026_05_02` claims:
  centering_lh = round(max(natural_lh, font_size × 83/64))

Spec §3.3's current formula (verified on TNR/MS Mincho/Yu/Meiryo at pitch=18):
  uses raw lm0_lh, no 83/64 floor.

Goal: test if 83/64 floor applies UNIVERSALLY (Latin too) or only conditionally
(e.g., single-cell case). Test cases:
  - Calibri 18pt at pitch=24: memory predicts offset=0.5 (with floor)
  - MSM 18pt at pitch=24: memory predicts offset=0.5
  - TNR 18pt at pitch=18: spec says offset=7.5 (no floor)
  - Calibri 18pt at pitch=18: NEW — distinguishes the two formulas
  - Calibri 11pt at pitch=24: NEW — larger pitch, single-cell

If single-cell-only floor:
  Calibri 18 pitch=18 (multi-cell, natural=22 > 18): no floor → offset=7.0
  Calibri 18 pitch=24 (single-cell, natural=22 < 24): floor → offset=0.5
  TNR 18 pitch=18 (multi-cell): no floor → offset=7.5
"""
import os
import sys
import time
import json

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "centering_floor_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_centering_lh_universal_floor.json")

WD_LINE_SPACE_SINGLE = 0
WD_LAYOUT_LINEGRID = 2

# Test cases: (font, font_size_pt, pitch_tw)
CASES = [
    # Memory's claimed cases (single-cell, natural < pitch)
    ("Calibri",          18.0, 480),  # pitch 24pt — memory says offset=0.5
    ("MS Mincho",        18.0, 480),  # pitch 24pt — memory says offset=0.5
    # Spec's existing case (multi-cell)
    ("Times New Roman",  18.0, 360),  # pitch 18pt — spec says offset=7.5
    # NEW distinguishing cases
    ("Calibri",          18.0, 360),  # pitch 18pt — single vs multi distinguishes
    ("Calibri",          11.0, 480),  # pitch 24pt — single-cell, Latin
    ("Calibri",          11.0, 360),  # pitch 18pt — control
    ("Times New Roman",  18.0, 480),  # pitch 24pt — Latin natural < pitch
    ("MS Mincho",        14.0, 480),  # pitch 24pt — single-cell CJK
    ("Yu Mincho",        18.0, 480),  # pitch 24pt — Yu Mincho natural ≈ 30 > pitch
]


def fareast_for(font):
    if font in ("Calibri", "Times New Roman"):
        return "MS Mincho"
    return font


def build_and_measure(word, font, size_pt, pitch_tw):
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.LayoutMode = WD_LAYOUT_LINEGRID
    ps.LinesPage = int(round(
        (ps.PageHeight - ps.TopMargin - ps.BottomMargin) * 20 / pitch_tw
    ))

    wdoc.Content.Text = ""
    rng = wdoc.Range(0, 0)
    rng.InsertAfter("Sample\r")
    rng.InsertAfter("Sample\r")
    rng.InsertAfter("Sample")

    for i in range(1, 4):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = font
        try:
            p.Range.Font.NameAscii = font
        except Exception:
            pass
        try:
            p.Range.Font.NameFarEast = fareast_for(font)
        except Exception:
            pass
        p.Range.Font.Size = size_pt
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0

    fname = f"CFLOOR_{font.replace(' ','')}_{size_pt}pt_p{pitch_tw}.docx"
    path = os.path.join(FIX_DIR, fname)
    wdoc.SaveAs2(path)
    wdoc.Close(False)

    wdoc = word.Documents.Open(path)
    try:
        wdoc.Repaginate()
        time.sleep(0.05)
        ys = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            ys.append(round(p.Range.Information(6), 4))
        return ys
    finally:
        wdoc.Close(False)


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)

    results = []
    try:
        for font, size_pt, pitch_tw in CASES:
            try:
                ys = build_and_measure(word, font, size_pt, pitch_tw)
                if len(ys) >= 3:
                    p1 = ys[0]
                    p2 = ys[1]
                    gap = round(p2 - p1, 4)
                    p1_offset = round(p1 - 72.0, 4)  # offset from top margin
                    pitch_pt = pitch_tw / 20.0
                    print(f"  {font[:14]:14s} {size_pt:>5.1f}pt p{pitch_tw}tw"
                          f"({pitch_pt}pt): p1_y={p1} offset={p1_offset:+.2f}pt"
                          f" gap={gap}pt")
                    results.append({
                        "font": font, "size_pt": size_pt,
                        "pitch_tw": pitch_tw, "pitch_pt": pitch_pt,
                        "ys": ys, "p1_offset_from_topmargin": p1_offset,
                        "gap_p2_p1": gap,
                    })
            except Exception as e:
                print(f"  ERR {font} {size_pt}pt p{pitch_tw}: {e}")
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved {len(results)} records.")
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
