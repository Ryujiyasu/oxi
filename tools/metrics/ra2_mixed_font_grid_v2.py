"""
Ra2: §1.7 mixed-font line height grid path — Word-native re-measurement.

Previous python-docx attempt (ra2_mixed_font.py) had Normal-style
Multiple 1.15 + sa=8pt defaults that contaminated grid measurements
(11/16 grid fixtures failed simple max rule).

This v2 builds fixtures via Word COM directly with explicit:
  - Single line spacing (LineSpacingRule = wdLineSpaceSingle = 0)
  - SpaceAfter = 0
  - SpaceBefore = 0
  - Explicit docGrid via Section.PageSetup.LayoutMode + LinesPage

Each fixture has 4 paragraphs:
  P1: pure font A
  P2: pure font B
  P3: MIXED line (run-A + run-B + run-A)
  P4: pure font A (control — gap from P3 = mixed line height)

Sweep: 4 font pairs × 4 size combos × 2 grid configs = 32 fixtures.

Goal: re-test simple max rule and determine if grid actually breaks it,
or if it was a python-docx-default artifact.
"""
import os
import time
import json
import sys

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "mixed_font_grid_v2_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_mixed_font_grid_v2.json")


WD_LINE_SPACE_SINGLE = 0
WD_LAYOUT_DEFAULT = 0
WD_LAYOUT_LINEGRID = 2


def add_run(para, text, font, size_pt, before_run=None):
    """Add a run with text/font/size, ensuring east-asia name set."""
    if before_run is None:
        run = para.Range.Document.Range(para.Range.End - 1, para.Range.End - 1)
    else:
        run = before_run
    run.Text = text
    run.Font.Name = font
    run.Font.Size = size_pt
    try:
        run.Font.NameFarEast = font
        run.Font.NameAscii = font
    except Exception:
        pass


def build_fixture(word, out_path, *, font_a, size_a, font_b, size_b,
                   pitch_tw=0):
    """Build a 4-paragraph fixture via Word COM."""
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

    # P1: pure font A
    # We build content via a sequence of inserts, then format each paragraph
    wdoc.Content.Text = ""
    rng = wdoc.Range(0, 0)
    rng.InsertAfter("AAA pure\r")
    rng.InsertAfter("BBB pure\r")
    # P3 mixed line — we'll set runs after; for now insert text
    rng.InsertAfter("MIX-A MIX-B back-A\r")
    rng.InsertAfter("AAA control")

    # Format P1
    p1 = wdoc.Paragraphs(1)
    p1.Range.Font.Name = font_a
    try:
        p1.Range.Font.NameFarEast = font_a
    except Exception:
        pass
    p1.Range.Font.Size = size_a
    p1.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    p1.Format.SpaceBefore = 0
    p1.Format.SpaceAfter = 0

    # Format P2
    p2 = wdoc.Paragraphs(2)
    p2.Range.Font.Name = font_b
    try:
        p2.Range.Font.NameFarEast = font_b
    except Exception:
        pass
    p2.Range.Font.Size = size_b
    p2.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    p2.Format.SpaceBefore = 0
    p2.Format.SpaceAfter = 0

    # P3 mixed: split text into 3 segments by character index
    p3 = wdoc.Paragraphs(3)
    p3.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    p3.Format.SpaceBefore = 0
    p3.Format.SpaceAfter = 0
    # Set baseline font A for whole P3
    p3.Range.Font.Name = font_a
    p3.Range.Font.Size = size_a
    try:
        p3.Range.Font.NameFarEast = font_a
    except Exception:
        pass
    # Override middle segment "MIX-B " (chars 6..12) with font B
    p3_text = p3.Range.Text  # includes trailing \r
    mix_b_start = p3_text.find("MIX-B")
    mix_b_end = mix_b_start + len("MIX-B ")
    if mix_b_start >= 0:
        sub = wdoc.Range(p3.Range.Start + mix_b_start,
                         p3.Range.Start + mix_b_end)
        sub.Font.Name = font_b
        sub.Font.Size = size_b
        try:
            sub.Font.NameFarEast = font_b
        except Exception:
            pass

    # Format P4
    p4 = wdoc.Paragraphs(4)
    p4.Range.Font.Name = font_a
    try:
        p4.Range.Font.NameFarEast = font_a
    except Exception:
        pass
    p4.Range.Font.Size = size_a
    p4.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    p4.Format.SpaceBefore = 0
    p4.Format.SpaceAfter = 0

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

    font_pairs = [
        ("Calibri",         "MS Mincho"),
        ("Times New Roman", "MS Gothic"),
        ("Yu Mincho",       "Yu Gothic"),
        ("Calibri",         "Yu Mincho"),
    ]
    size_combos = [(8, 8), (11, 14), (14, 11), (18, 24)]
    grid_options = [(0, "noGrid"), (360, "grid360tw")]

    results = []
    try:
        for font_a, font_b in font_pairs:
            for size_a, size_b in size_combos:
                for pitch, glabel in grid_options:
                    fname = f"MFG2_{font_a.replace(' ','')}{size_a}_{font_b.replace(' ','')}{size_b}_{glabel}.docx"
                    path = os.path.join(FIX_DIR, fname)
                    try:
                        build_fixture(word, path,
                                      font_a=font_a, size_a=size_a,
                                      font_b=font_b, size_b=size_b,
                                      pitch_tw=pitch)
                        ys = measure_fixture(word, path)
                        if len(ys) >= 4:
                            g_p1p2 = round(ys[1] - ys[0], 4)
                            g_p2p3 = round(ys[2] - ys[1], 4)
                            g_p3p4 = round(ys[3] - ys[2], 4)
                            expected_max = max(g_p1p2, g_p2p3)
                            match = abs(g_p3p4 - expected_max) < 0.6
                            mark = "OK" if match else "FAIL"
                            print(f"  {font_a[:8]:8s}{size_a}/{font_b[:8]:8s}{size_b} ({glabel}): "
                                  f"A={g_p1p2} B={g_p2p3} mix={g_p3p4} max={expected_max} {mark}")
                            results.append({
                                "font_a": font_a, "size_a": size_a,
                                "font_b": font_b, "size_b": size_b,
                                "pitch_tw": pitch, "grid_label": glabel,
                                "ys": ys,
                                "gap_a": g_p1p2, "gap_b": g_p2p3, "gap_mix": g_p3p4,
                                "expected_max": expected_max,
                                "match": match,
                            })
                    except Exception as e:
                        print(f"  ERR {fname}: {e}")
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
