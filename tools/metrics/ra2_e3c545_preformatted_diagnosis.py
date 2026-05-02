"""e3c545 PreformattedText 9pt — diagnose the real drift cause.

Hypotheses to test:
1. ASCII char widths: does Word use MS PGothic proportional widths for ASCII
   chars in a PreformattedText style (ascii="ＭＳ Ｐゴシック")?
2. Per-line LH: pure MS PGothic 9pt empty/short paragraph line height —
   verify it matches Oxi's 11.625pt (floor formula).
3. Empty-paragraph height in PreformattedText style — does it differ from
   regular empty paragraph?

Three test docs:
  T1: 5 short PreformattedText paragraphs ("Pure" each) — line gap = LH per
      paragraph (no wrap, simple gap).
  T2: PreformattedText with mixed ASCII+CJK content "@prefix rdf:に" — measure
      per-char widths.
  T3: 3 PreformattedText paragraphs alternating short/empty/short — measure
      empty para height vs short para height.
"""
import os
import sys
import time
import json

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "e3c545_preformatted_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_e3c545_preformatted_diagnosis.json")

WD_LINE_SPACE_SINGLE = 0
WD_LAYOUT_DEFAULT = 0


def setup_pre_para(p, font, size, fareast):
    p.Range.Font.Name = font
    try:
        p.Range.Font.NameAscii = font
    except Exception:
        pass
    try:
        p.Range.Font.NameFarEast = fareast
    except Exception:
        pass
    p.Range.Font.Size = size
    p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    p.Format.SpaceBefore = 0
    p.Format.SpaceAfter = 0


def build_t1(word, path, font="ＭＳ Ｐゴシック", size=9.0):
    """T1: 5 short paragraphs to measure line gap."""
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.LayoutMode = WD_LAYOUT_DEFAULT  # noGrid
    wdoc.Content.Text = ""
    rng = wdoc.Range(0, 0)
    for i in range(4):
        rng.InsertAfter(f"Pure{i+1}\r")
    rng.InsertAfter("Pure5")
    for i in range(1, 6):
        setup_pre_para(wdoc.Paragraphs(i), font, size, font)
    wdoc.SaveAs2(path)
    wdoc.Close(False)


def build_t2(word, path, font="ＭＳ Ｐゴシック", size=9.0):
    """T2: per-char width sweep on @prefix rdf:Hello (mixed ASCII)."""
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.LayoutMode = WD_LAYOUT_DEFAULT
    wdoc.Content.Text = ""
    rng = wdoc.Range(0, 0)
    rng.InsertAfter("@prefix rdf:Hello")
    setup_pre_para(wdoc.Paragraphs(1), font, size, font)
    wdoc.SaveAs2(path)
    wdoc.Close(False)


def build_t3(word, path, font="ＭＳ Ｐゴシック", size=9.0):
    """T3: empty paragraph between short paragraphs."""
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.LayoutMode = WD_LAYOUT_DEFAULT
    wdoc.Content.Text = ""
    rng = wdoc.Range(0, 0)
    rng.InsertAfter("Top\r")    # P1
    rng.InsertAfter("\r")        # P2 empty
    rng.InsertAfter("\r")        # P3 empty
    rng.InsertAfter("Bot")       # P4
    for i in range(1, 5):
        setup_pre_para(wdoc.Paragraphs(i), font, size, font)
    wdoc.SaveAs2(path)
    wdoc.Close(False)


def measure_paragraph_ys(word, path):
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


def measure_per_char_x(word, path):
    """Returns list of (char, x, y) for first paragraph, char by char."""
    wdoc = word.Documents.Open(path)
    try:
        wdoc.Repaginate()
        time.sleep(0.05)
        para = wdoc.Paragraphs(1).Range
        results = []
        for i in range(para.Start, para.End):
            sub = wdoc.Range(i, i + 1)
            ch = sub.Text
            x = round(sub.Information(5), 4)
            y = round(sub.Information(6), 4)
            results.append({"char": ch, "x": x, "y": y})
        return results
    finally:
        wdoc.Close(False)


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)

    results = {}
    try:
        # T1: short-paragraph line gap
        path1 = os.path.join(FIX_DIR, "T1_pre_short_paras.docx")
        build_t1(word, path1)
        ys = measure_paragraph_ys(word, path1)
        gaps = [round(ys[i+1] - ys[i], 4) for i in range(len(ys) - 1)]
        results["T1_short_paras"] = {"ys": ys, "gaps": gaps}
        print(f"T1 ys: {ys}")
        print(f"T1 gaps: {gaps}")

        # T2: per-char width
        path2 = os.path.join(FIX_DIR, "T2_per_char.docx")
        build_t2(word, path2)
        chars = measure_per_char_x(word, path2)
        results["T2_per_char_widths"] = chars
        print(f"\nT2 per-char advances:")
        for c in chars:
            ch = c["char"]
            x = c["x"]
            print(f"  {ch!r:>6s}  x={x}")

        # T3: empty para height
        path3 = os.path.join(FIX_DIR, "T3_empty_para.docx")
        build_t3(word, path3)
        ys = measure_paragraph_ys(word, path3)
        gaps = [round(ys[i+1] - ys[i], 4) for i in range(len(ys) - 1)]
        results["T3_empty_para"] = {"ys": ys, "gaps": gaps}
        print(f"\nT3 ys: {ys}")
        print(f"T3 gaps: {gaps}")
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved to {OUT_JSON}")
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
