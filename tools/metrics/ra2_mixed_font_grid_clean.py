"""
Ra2: §1.7 grid mixed-font line height — CLEAN re-measurement with isolated gaps.

v2 measurement was confused: it interpreted P2-P1 gap (Calibri 11 → MS Mincho 14)
as "pure Calibri 11 line height", but it's actually a cross-paragraph distance
that depends on Word's grid alignment for BOTH adjacent paragraphs.

This tool isolates each line height with same-font paragraph pairs:
  P1, P2: pure font_a (gap P2-P1 = pure font_a line height)
  P3, P4: pure font_b (gap P4-P3 = pure font_b line height)
  P5, P6: mixed font_a + font_b run (gap P6-P5 = mixed line height)

After running, gap_pure_a / gap_pure_b / gap_mix give the three needed values
with no cross-contamination from grid alignment of differently-sized adjacent
paragraphs.
"""
import os
import time
import json
import sys

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "mixed_font_clean_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_mixed_font_grid_clean.json")

WD_LINE_SPACE_SINGLE = 0
WD_LAYOUT_DEFAULT = 0
WD_LAYOUT_LINEGRID = 2

DEFAULT_CJK_FALLBACK = "MS Mincho"


def is_cjk_font(name: str) -> bool:
    return any(m in name for m in ("Mincho", "Gothic", "Meiryo", "明朝",
                                    "ゴシック", "メイリオ", "HG"))


def fareast_for(font: str) -> str:
    return font if is_cjk_font(font) else DEFAULT_CJK_FALLBACK


def set_para_font(p, font, size):
    p.Range.Font.Name = font
    try:
        p.Range.Font.NameAscii = font
    except Exception:
        pass
    try:
        p.Range.Font.NameFarEast = fareast_for(font)
    except Exception:
        pass
    p.Range.Font.Size = size
    p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    p.Format.SpaceBefore = 0
    p.Format.SpaceAfter = 0


def build_fixture(word, out_path, *, font_a, size_a, font_b, size_b, pitch_tw):
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
    rng.InsertAfter("AAA1\r")              # P1 pure-A
    rng.InsertAfter("AAA2\r")              # P2 pure-A
    rng.InsertAfter("BBB1\r")              # P3 pure-B
    rng.InsertAfter("BBB2\r")              # P4 pure-B
    rng.InsertAfter("MIX1A MIX1B back\r")  # P5 mixed
    rng.InsertAfter("MIX2A MIX2B back")    # P6 mixed

    # P1, P2 pure font A
    for i in (1, 2):
        set_para_font(wdoc.Paragraphs(i), font_a, size_a)

    # P3, P4 pure font B
    for i in (3, 4):
        set_para_font(wdoc.Paragraphs(i), font_b, size_b)

    # P5, P6 mixed: baseline font A, override middle "MIXxB " segment with font B
    for i in (5, 6):
        p = wdoc.Paragraphs(i)
        set_para_font(p, font_a, size_a)
        p_text = p.Range.Text
        marker = f"MIX{i-4}B"
        idx = p_text.find(marker)
        if idx >= 0:
            sub = wdoc.Range(p.Range.Start + idx,
                             p.Range.Start + idx + len(marker) + 1)
            sub.Font.Name = font_b
            try:
                sub.Font.NameAscii = font_b
            except Exception:
                pass
            try:
                sub.Font.NameFarEast = fareast_for(font_b)
            except Exception:
                pass
            sub.Font.Size = size_b

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

    pairs = [
        ("Calibri",         "MS Mincho"),
        ("Times New Roman", "MS Gothic"),
        ("Yu Mincho",       "Yu Gothic"),
        ("Calibri",         "Yu Mincho"),
    ]
    size_combos = [(8, 8), (11, 14), (14, 11), (18, 24)]
    grids = [(0, "noGrid"), (320, "g320"), (360, "g360"), (440, "g440")]

    results = []
    fail_count = 0
    try:
        for font_a, font_b in pairs:
            for size_a, size_b in size_combos:
                for pitch_tw, glabel in grids:
                    fname = (
                        f"MFC_{font_a.replace(' ','')}{size_a}_"
                        f"{font_b.replace(' ','')}{size_b}_{glabel}.docx"
                    )
                    path = os.path.join(FIX_DIR, fname)
                    try:
                        build_fixture(word, path,
                                      font_a=font_a, size_a=size_a,
                                      font_b=font_b, size_b=size_b,
                                      pitch_tw=pitch_tw)
                        ys = measure_fixture(word, path)
                        if len(ys) >= 6:
                            gap_a = round(ys[1] - ys[0], 4)
                            gap_b = round(ys[3] - ys[2], 4)
                            gap_mix = round(ys[5] - ys[4], 4)
                            results.append({
                                "font_a": font_a, "size_a": size_a,
                                "font_b": font_b, "size_b": size_b,
                                "pitch_tw": pitch_tw, "grid_label": glabel,
                                "ys": ys,
                                "gap_pure_a": gap_a,
                                "gap_pure_b": gap_b,
                                "gap_mix": gap_mix,
                            })
                            print(f"  {font_a[:8]:8s}{size_a}/{font_b[:8]:8s}{size_b} {glabel}:"
                                  f" pure_a={gap_a} pure_b={gap_b} mix={gap_mix}")
                    except Exception as e:
                        fail_count += 1
                        print(f"  ERR {fname}: {e}")
                        if fail_count >= 5:
                            print("  Restarting Word...")
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
