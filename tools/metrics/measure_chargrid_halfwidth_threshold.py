"""§11.2.1 charGrid half-width snap threshold sweep — robust v2.

Strategy: ONE Word session, ONE doc per (font, size, charsLine) combo, ONE
paragraph holding all probe chars inline (separated by 一 to keep them
flanked by full-width). Reduces COM call count to minimum.

Also: measure inline probe positions via Range character iteration without
TypeText/Selection.
"""
import win32com.client
import os
import time
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Probe chars span natural-width / fontSize ratios from ~0.2 to 1.0
PROBE_CHARS = [
    ("kana_a",       "あ"),
    ("ideo_one",     "一"),
    ("kata_a",       "ア"),
    ("zen_space",    "　"),
    ("fw_A",         "Ａ"),
    ("Latin_M",      "M"),
    ("Latin_W",      "W"),
    ("Latin_A",      "A"),
    ("Latin_i",      "i"),
    ("Latin_l",      "l"),
    ("Latin_0",      "0"),
    ("dot",          "."),
    ("excl",         "!"),
    ("hw_kata_a",    "ｱ"),
]

# Build a single text where each probe is bracketed by 一 so we can identify
# probe position robustly: the i-th probe is at character index 2*i+1 (0-based).
# i.e., the structure is 一 P0 一 P1 一 P2 一 ...
def build_probe_text():
    chars = []
    for _, ch in PROBE_CHARS:
        chars.append("一")
        chars.append(ch)
    chars.append("一")  # tail
    return "".join(chars)


SIZE_CHARSLINE = [
    (10.5, 40),  # pitch ≈ 10.6pt (~ fontSize)
    (12.0, 35),  # pitch ≈ 12.1pt
    (14.0, 30),  # pitch ≈ 14.1pt
    (10.5, 30),  # pitch ≈ 14.2pt (>> fontSize)
]

FONTS = ["ＭＳ ゴシック", "ＭＳ 明朝", "Yu Gothic", "Meiryo"]
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    probe_text = build_probe_text()
    try:
        for font in FONTS:
            results[font] = {}
            for fontSize, charsLine in SIZE_CHARSLINE:
                key = f"sz{fontSize}_chars{charsLine}"
                results[font][key] = {"fontSize": fontSize,
                                      "charsLine": charsLine,
                                      "chars": {}}
                try:
                    doc = word.Documents.Add()
                    time.sleep(0.5)
                    sec = doc.Sections(1)
                    ps = sec.PageSetup
                    ps.PageHeight = 841.9
                    ps.PageWidth = 595.3
                    ps.TopMargin = 72
                    ps.BottomMargin = 72
                    ps.LeftMargin = 85
                    ps.RightMargin = 85
                    ps.LayoutMode = 2  # wdLayoutModeGenko
                    ps.CharsLine = charsLine
                    content_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
                    pitch = content_w / charsLine
                    results[font][key]["content_w"] = round(content_w, 4)
                    results[font][key]["pitch"] = round(pitch, 4)
                    # Set body text
                    rng = doc.Range()
                    rng.Text = probe_text
                    rng = doc.Range()
                    rng.Font.Name = font
                    rng.Font.Size = fontSize
                    para = doc.Paragraphs(1)
                    para.Format.LeftIndent = 0
                    para.Format.FirstLineIndent = 0
                    para.Format.SpaceBefore = 0
                    para.Format.SpaceAfter = 0
                    time.sleep(0.2)
                    # Measure each character's x via Information(5)
                    chars = doc.Range().Characters
                    xs = []
                    for ci in range(1, chars.Count + 1):
                        try:
                            c = chars(ci)
                            t = c.Text
                            if t in ("\r", "\x07"):
                                continue
                            xs.append((t, float(c.Information(5))))
                        except Exception:
                            continue
                    # Each probe at index 2i+1, prev/next 一 at 2i and 2i+2
                    # advance(probe_i) = xs[2i+2].x - xs[2i+1].x
                    for i, (label, ch) in enumerate(PROBE_CHARS):
                        idx_probe = 2 * i + 1
                        idx_after = 2 * i + 2
                        if idx_after < len(xs) and xs[idx_probe][0] == ch:
                            adv = round(xs[idx_after][1] - xs[idx_probe][1], 4)
                            cjk_adv = (round(xs[idx_probe][1]
                                              - xs[idx_probe - 1][1], 4)
                                        if idx_probe - 1 >= 0 else None)
                        else:
                            adv = None
                            cjk_adv = None
                        results[font][key]["chars"][label] = {
                            "char": ch,
                            "ord": ord(ch),
                            "probe_advance": adv,
                            "cjk_advance": cjk_adv,
                            "ratio_to_pitch": (round(adv / pitch, 4)
                                                if adv else None),
                            "ratio_to_size": (round(adv / fontSize, 4)
                                               if adv else None),
                        }
                        line = (f"[{font}][{key}][{label:14s}] "
                                f"adv={adv} pitch={pitch:.3f} "
                                f"r/pitch={(adv/pitch if adv else 0):.3f} "
                                f"r/size={(adv/fontSize if adv else 0):.3f}")
                        print(line, flush=True)
                    doc.Close(SaveChanges=False)
                except Exception as e:
                    results[font][key]["error"] = str(e)
                    print(f"[{font}][{key}] FATAL: {e}", flush=True)
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    if os.path.exists(RESULT_PATH):
        try:
            with open(RESULT_PATH, encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = {}
    else:
        existing = {}
    existing["chargrid_halfwidth_threshold_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
