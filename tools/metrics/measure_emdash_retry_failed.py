"""§4.7 em-dash retry — fonts that failed in main sweep.

Yu Gothic, HGS明朝E, HG明朝B, HGゴシックE 14/18 all hit RPC errors in the
main sweep. This script retries them ONE FONT AT A TIME with full Word
restart between sizes.
"""
import win32com.client
import os
import time
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONTS = ["Yu Gothic", "HGS明朝E", "HG明朝B", "HGゴシックE"]
SIZES = [10.5, 12.0, 14.0, 18.0]
PROBE_CHARS = [
    ("emdash",     "—"),
    ("hbar",       "―"),
    ("ckakko",     "」"),
    ("toten",      "、"),
]
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def measure_one(font, size, label, ch):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        text = f"漢{ch}{ch}漢"
        d = word.Documents.Add()
        time.sleep(0.5)
        rng = d.Range()
        rng.InsertAfter(text)
        rng = d.Range()
        rng.Font.Name = font
        rng.Font.Size = size
        d.Paragraphs(1).Alignment = 0
        time.sleep(0.2)
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
        return [(xs[i][0], round(xs[i + 1][1] - xs[i][1], 4))
                for i in range(len(xs) - 1)]
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        time.sleep(1.0)


def main():
    if os.path.exists(RESULT_PATH):
        with open(RESULT_PATH, encoding="utf-8") as f:
            existing = json.load(f)
    else:
        existing = {}
    block = existing.get("emdash_font_sweep_2026-05-02", {})

    for font in FONTS:
        if font not in block:
            block[font] = {}
        for size in SIZES:
            ks = str(size)
            if ks not in block[font]:
                block[font][ks] = {}
            for label, ch in PROBE_CHARS:
                # Skip if already have valid data
                if label in block[font][ks]:
                    advs = block[font][ks][label].get("advances")
                    if isinstance(advs, list) and len(advs) >= 3:
                        continue
                try:
                    advs = measure_one(font, size, label, ch)
                except Exception as e:
                    advs = {"error": str(e)}
                block[font][ks][label] = {
                    "text": f"漢{ch}{ch}漢",
                    "advances": advs,
                }
                if isinstance(advs, list) and len(advs) >= 2:
                    first_X = advs[1][1]
                    ratio = round(first_X / size, 3)
                    marker = " <-- COMPRESSED" if ratio < 0.6 else " (full)"
                    print(f"[{font}][{size}][{label}] adv1st={first_X} "
                          f"r={ratio}{marker}", flush=True)
                else:
                    print(f"[{font}][{size}][{label}] err: {advs}",
                          flush=True)

    existing["emdash_font_sweep_2026-05-02"] = block
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
