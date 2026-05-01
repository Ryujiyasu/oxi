"""§4.7 em-dash compression font-dependence sweep.

Earlier finding (4-font sweep): em-dash (U+2014) compresses in MS 明朝
and MS ゴシック (Type B) but NOT in Yu Mincho (acts like Type C).
Tested only Yu Mincho 10.5pt — broader sweep needed.

This script tests across:
- 5+ fonts (MS 明朝, MS ゴシック, Yu Mincho, Yu Gothic, Meiryo, HG variants)
- 4 sizes (10.5, 12, 14, 18)
- 4 Type-B chars: — (U+2014), ― (U+2015 control), 」 (universal B), ） (universal B)

For each (font, size, char), measure:
  - 漢XX漢 (paired: 1st X should compress per FINAL RULE if char is B-class
    in this font)

Builds a "font-by-char compression table" replacing the spec's font-agnostic
classification.
"""
import win32com.client
import os
import time
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONTS = [
    "ＭＳ 明朝",
    "ＭＳ ゴシック",
    "Yu Mincho",
    "Yu Gothic",
    "Meiryo",
    "HGS明朝E",
    "HGゴシックE",
    "HG明朝B",
]
SIZES = [10.5, 12.0, 14.0, 18.0]
PROBE_CHARS = [
    ("emdash",     "—"),   # U+2014
    ("hbar",       "―"),   # U+2015
    ("ckakko",     "」"),
    ("cparen",     "）"),
    ("toten",      "、"),
    ("kuten",      "。"),
]

RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def make_word():
    w = win32com.client.Dispatch("Word.Application")
    w.Visible = False
    w.DisplayAlerts = False
    return w


def main():
    results = {}
    # Restart Word per font to dodge RPC death
    for font in FONTS:
        results[font] = {}
        word = make_word()
        try:
            for size in SIZES:
                results[font][str(size)] = {}
                for label, ch in PROBE_CHARS:
                    text = f"漢{ch}{ch}漢"  # paired probe: B→B trigger
                    try:
                        d = word.Documents.Add()
                        time.sleep(0.2)
                        rng = d.Range()
                        rng.InsertAfter(text)
                        rng = d.Range()
                        rng.Font.Name = font
                        rng.Font.Size = size
                        d.Paragraphs(1).Alignment = 0
                        time.sleep(0.1)
                        chars = d.Range().Characters
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
                        d.Close(SaveChanges=False)
                        advs = [(xs[i][0],
                                 round(xs[i + 1][1] - xs[i][1], 4))
                                for i in range(len(xs) - 1)]
                    except Exception as e:
                        advs = {"error": str(e)}
                    results[font][str(size)][label] = {
                        "text": text,
                        "advances": advs,
                    }
                    # Compute first-X advance ratio to fontSize for compress
                    # detection
                    if isinstance(advs, list) and len(advs) >= 2:
                        first_X_adv = advs[1][1]  # second char in pair
                        ratio = round(first_X_adv / size, 3)
                        compress_marker = ""
                        if ratio < 0.6:
                            compress_marker = " <-- COMPRESSED"
                        elif ratio > 0.9:
                            compress_marker = " (full width)"
                        line = (f"[{font}][{size}][{label:8s}] "
                                f"adv1st={first_X_adv} r={ratio}"
                                + compress_marker)
                        print(line, flush=True)
                    else:
                        print(f"[{font}][{size}][{label}] err: {advs}",
                              flush=True)
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            time.sleep(1.0)

    if os.path.exists(RESULT_PATH):
        try:
            with open(RESULT_PATH, encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = {}
    else:
        existing = {}
    existing["emdash_font_sweep_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
