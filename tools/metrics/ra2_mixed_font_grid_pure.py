"""
Ra2: §1.7 grid mode pure-paragraph line height — characterize per-font, per-size.

Background:
  ra2_mixed_font_grid_v2.py reported gap_a (Calibri 11pt, single LH, grid 18pt)
  = 25.0pt. Expected ~18pt (1 grid cell) since Calibri 11 natural lh ≈ 13.5pt.
  Hypothesis: v2 sets `NameFarEast = font_a` for Latin font_a (e.g., Calibri),
  which causes Word to substitute a default CJK font for the paragraph mark
  with mismatched metrics, inflating the line height.

This tool:
  1. Tests pure-font 1-line paragraph height across (font, size, pitch).
  2. Compares two NameFarEast strategies:
       (A) NameFarEast = same as Latin font (the v2 way — likely buggy)
       (B) NameFarEast = MS Mincho (proper CJK fallback for Latin paragraphs)
  3. For CJK fonts (MS Mincho, Yu Mincho), NameFarEast = same as font (correct).

Output format:
  JSON list of records with `font`, `size`, `pitch_tw`, `farEast_strategy`,
  `ys` (paragraph Y positions for 4 paragraphs), `gap` (the line height).

After running, analyze: does strategy (B) produce gap matching
`ceil(natural_lh / pitch) * pitch` (for natural<pitch) or proportional?
Strategy (A) reproducing v2's anomaly would confirm the hypothesis.
"""
import os
import time
import json
import sys

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "mixed_font_pure_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_mixed_font_grid_pure.json")

WD_LINE_SPACE_SINGLE = 0
WD_LAYOUT_DEFAULT = 0
WD_LAYOUT_LINEGRID = 2

# Latin-font default CJK fallback we will use for strategy B.
DEFAULT_CJK_FALLBACK = "MS Mincho"


def is_cjk_font(name: str) -> bool:
    """Heuristic — is this a font that has CJK glyphs natively?"""
    cjk_markers = ("Mincho", "Gothic", "Meiryo", "明朝", "ゴシック", "メイリオ", "HG")
    return any(m in name for m in cjk_markers)


def build_fixture(word, out_path, *, font, size, pitch_tw, fareast_strategy):
    """4-paragraph fixture, all paragraphs identical pure-font.
    P1..P4 = "AAA pure" with given font+size, Single LH, sa=sb=0.
    P2-P1 gap = single-line paragraph height under those settings.
    """
    if is_cjk_font(font):
        fareast = font
    elif fareast_strategy == "A_same":
        fareast = font  # v2 way (Latin as FarEast — the buggy way)
    elif fareast_strategy == "B_msmincho":
        fareast = DEFAULT_CJK_FALLBACK
    else:
        raise ValueError(fareast_strategy)

    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.HeaderDistance = 36
    ps.FooterDistance = 36
    if pitch_tw == 0:
        ps.LayoutMode = WD_LAYOUT_DEFAULT
    else:
        ps.LayoutMode = WD_LAYOUT_LINEGRID
        ps.LinesPage = int(round(
            (ps.PageHeight - ps.TopMargin - ps.BottomMargin) * 20 / pitch_tw
        ))

    wdoc.Content.Text = ""
    rng = wdoc.Range(0, 0)
    rng.InsertAfter("Pure 1\r")
    rng.InsertAfter("Pure 2\r")
    rng.InsertAfter("Pure 3\r")
    rng.InsertAfter("Pure 4")

    for i in range(1, 5):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = font
        try:
            p.Range.Font.NameFarEast = fareast
        except Exception:
            pass
        try:
            p.Range.Font.NameAscii = font
        except Exception:
            pass
        p.Range.Font.Size = size
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0

    wdoc.SaveAs2(out_path)
    wdoc.Close(False)


def measure_fixture(word, path):
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

    # Sweep matrix: small but pinpointed at the v2 anomaly.
    # font, sizes
    fonts = [
        ("Calibri",          [8, 11, 14, 18]),
        ("Times New Roman",  [8, 11, 14, 18]),
        ("MS Mincho",        [8, 11, 14, 18]),
        ("MS Gothic",        [8, 11, 14, 18]),
        ("Yu Mincho",        [11, 14, 18]),
        ("Meiryo",           [11, 14]),
    ]
    grids = [(0, "noGrid"), (240, "g240"), (320, "g320"),
             (360, "g360"), (440, "g440")]
    fareast_strategies = ["A_same", "B_msmincho"]

    results = []
    fail_count = 0
    try:
        for font, sizes in fonts:
            for size in sizes:
                for pitch_tw, glabel in grids:
                    for strat in fareast_strategies:
                        # For CJK fonts both strategies collapse — only run once.
                        if is_cjk_font(font) and strat == "B_msmincho":
                            continue
                        fname = (
                            f"PURE_{font.replace(' ','')}{size}_{glabel}"
                            f"_{strat}.docx"
                        )
                        path = os.path.join(FIX_DIR, fname)
                        try:
                            build_fixture(word, path,
                                          font=font, size=size,
                                          pitch_tw=pitch_tw,
                                          fareast_strategy=strat)
                            ys = measure_fixture(word, path)
                            if len(ys) >= 4:
                                gaps = [round(ys[i+1] - ys[i], 4) for i in range(3)]
                                gap = gaps[1]  # P3-P2 (avoid first-para edge effects)
                                row = {
                                    "font": font, "size": size,
                                    "pitch_tw": pitch_tw, "grid_label": glabel,
                                    "fareast_strategy": strat,
                                    "ys": ys, "gaps": gaps, "gap": gap,
                                }
                                results.append(row)
                                tag = "CJK" if is_cjk_font(font) else strat
                                print(f"  {font[:14]:14s} {size:>3d}pt {glabel:>7s} {tag:>11s}: gaps={gaps} gap={gap}")
                        except Exception as e:
                            fail_count += 1
                            print(f"  ERR {fname}: {e}")
                            if fail_count >= 5:
                                print("  Too many failures — restarting Word...")
                                try:
                                    word.Quit()
                                except Exception:
                                    pass
                                time.sleep(2)
                                word = win32com.client.gencache.EnsureDispatch("Word.Application")
                                word.Visible = False
                                word.DisplayAlerts = False
                                fail_count = 0
                                time.sleep(1.0)
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved {len(results)} records to {OUT_JSON}")
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
