"""
Ra2: Body start Y offset investigation (§8.2 word_layout_spec_ra.md).

Goal:
  Identify what determines `body_y - topMargin` for the first paragraph on a page.
  Spec currently says "topMargin + ~2.5pt offset" -- pin down the formula.

Variables swept:
  - top_margin (36, 72, 108, 144 pt)
  - header_distance (18, 36, 54 pt)
  - body font (Calibri, MS Gothic, MS Mincho, Yu Mincho, Times New Roman)
  - body font size (8, 10.5, 11, 14, 18 pt)
  - docGrid: noGrid vs lines(linePitch=360tw=18pt) vs lines(linePitch=480tw=24pt)

For each combo we record:
  body_y                = COM Information(6) of P1
  delta                 = body_y - top_margin
  natural_lh            = computed from win_ascent+win_descent (NOT measured here)
  ppem                  = round(font_size * 96 / 72)

Caveat (per MEMORY.md/com_information6_caveat):
  Information(6) = line box top, NOT glyph top.
"""
import win32com.client
import json
import os
import sys

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_body_start_offset.json")

FONTS = [
    ("Calibri", False),
    ("MS Gothic", True),
    ("MS Mincho", True),
    ("Yu Mincho", True),
    ("Times New Roman", False),
]
SIZES = [8, 10.5, 11, 14, 18]
TOP_MARGINS = [36, 72, 108]
HEADER_DISTS = [18, 36]
GRID_CONFIGS = [
    ("noGrid", 0, 0),       # docGrid not set
    ("lines360", 1, 360),   # type=lines, linePitch=360tw (18pt)
    ("lines480", 1, 480),   # type=lines, linePitch=480tw (24pt)
]

def make_doc(word, top_margin, header_dist, font, size, grid_type, line_pitch):
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = top_margin
    ps.BottomMargin = 72
    ps.HeaderDistance = header_dist
    ps.FooterDistance = 36

    # docGrid via XML manipulation is complex; use Word's own grid API.
    # Word.Document.Sections(1).PageSetup has no docGrid.
    # Use Document.NoLineBreakBefore? No -- use the LayoutMode property.
    # 0=wdLayoutModeDefault, 1=wdLayoutModeGrid, 2=wdLayoutModeLineGrid, 3=wdLayoutModeGenko
    if grid_type == 0:
        ps.LayoutMode = 0  # wdLayoutModeDefault (no grid)
    else:
        ps.LayoutMode = 2  # wdLayoutModeLineGrid (lines only)
        ps.LinesPage = int(round((ps.PageHeight - ps.TopMargin - ps.BottomMargin) * 20 / line_pitch))

    # Body
    wdoc.Content.Text = "本文1\r本文2\r本文3"
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = font
        # MS Gothic etc need EastAsia name as well
        try:
            p.Range.Font.NameFarEast = font
        except Exception:
            pass
        p.Range.Font.Size = size
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0  # wdLineSpaceSingle

    wdoc.Repaginate()

    body_p1 = wdoc.Paragraphs(1)
    body_p2 = wdoc.Paragraphs(2)
    p1_y = round(body_p1.Range.Information(6), 4)
    p2_y = round(body_p2.Range.Information(6), 4)

    rec = {
        "top_margin": top_margin,
        "header_dist": header_dist,
        "font": font,
        "size": size,
        "grid": ["noGrid", "lines360", "lines480"][grid_type if grid_type == 0 else (1 if line_pitch == 360 else 2)],
        "line_pitch_tw": line_pitch,
        "body_p1_y": p1_y,
        "body_p2_y": p2_y,
        "delta_p1_topmargin": round(p1_y - top_margin, 4),
        "p2_minus_p1": round(p2_y - p1_y, 4),
    }

    wdoc.Close(False)
    return rec


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    results = []
    try:
        # Phase A: vary topMargin x font x size, keep grid=noGrid, hdrDist=36
        print("=== Phase A: topMargin x font x size, noGrid ===")
        for tm in TOP_MARGINS:
            for font, _is_cjk in FONTS:
                for sz in SIZES:
                    try:
                        r = make_doc(word, tm, 36, font, sz, 0, 0)
                        results.append(r)
                        print(f"  tm={tm} {font:18s} {sz:5}pt -> p1_y={r['body_p1_y']:8.3f} delta={r['delta_p1_topmargin']:+6.3f} p2-p1={r['p2_minus_p1']:6.3f}")
                    except Exception as e:
                        print(f"  ERR tm={tm} {font} {sz}: {e}")

        # Phase B: vary headerDistance, fixed font/size/topMargin
        print("\n=== Phase B: headerDistance sweep, Calibri 11pt, tm=72 ===")
        for hd in HEADER_DISTS + [54, 6]:
            try:
                r = make_doc(word, 72, hd, "Calibri", 11, 0, 0)
                results.append(r)
                print(f"  hd={hd:3} -> p1_y={r['body_p1_y']:8.3f} delta={r['delta_p1_topmargin']:+6.3f}")
            except Exception as e:
                print(f"  ERR hd={hd}: {e}")

        # Phase C: grid variations
        print("\n=== Phase C: docGrid sweep, MS Gothic 10.5pt, tm=72 ===")
        for gt, _, lp in GRID_CONFIGS:
            try:
                gtype = 0 if gt == "noGrid" else 1
                r = make_doc(word, 72, 36, "MS Gothic", 10.5, gtype, lp)
                results.append(r)
                print(f"  {gt:10s} pitch={lp:4}tw -> p1_y={r['body_p1_y']:8.3f} delta={r['delta_p1_topmargin']:+6.3f} p2-p1={r['p2_minus_p1']:6.3f}")
            except Exception as e:
                print(f"  ERR {gt}: {e}")

    finally:
        word.Quit()

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved {len(results)} records to {OUT_JSON}")


if __name__ == "__main__":
    main()
