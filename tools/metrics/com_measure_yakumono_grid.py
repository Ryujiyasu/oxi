"""Session 113 — COM-measure punctuation advance for the parametric grid.

Outputs: tools/metrics/yakumono_grid/results.json
  Per variant: {font, fs, cs_tw, punct, dot_advance_pt, digit_advance_pt,
                kanji_advance_pt (mean of '提供を受けた匿名' inner CJK)}
"""
import os
import sys
import io
import re
import json
import glob
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
VARIANTS_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/yakumono_grid/variants"))
OUT_JSON = os.path.normpath(os.path.join(REPO, "tools/metrics/yakumono_grid/results.json"))

wdHorizontal = 5
wdVertical = 6

FONT_SLUG_TO_NAME = {
    "msmincho": "ＭＳ 明朝",
    "msgothic": "ＭＳ ゴシック",
    "meiryo": "メイリオ",
    "yumincho": "游明朝",
    "yugothic": "游ゴシック",
}
PUNCT_SLUG_TO_CHAR = {"dotF": '．', "commaF": '，', "kuten": '。'}


def parse_name(name):
    """Parse 'g_<fontslug>_sz<sz>_cs<cs>_<punctslug>' → meta dict."""
    # sz can be 2 digits; cs can be negative
    m = re.match(r"g_([a-z]+)_sz(\d+)_cs(-?\d+)_(\w+)", name)
    if not m:
        return None
    return {
        "font_slug": m.group(1),
        "font": FONT_SLUG_TO_NAME.get(m.group(1), m.group(1)),
        "sz_hp": int(m.group(2)),
        "fs_pt": int(m.group(2)) / 2.0,
        "cs_tw": int(m.group(3)),
        "punct": PUNCT_SLUG_TO_CHAR.get(m.group(4), m.group(4)),
    }


def measure_doc(word, path):
    doc = word.Documents.Open(path, ReadOnly=True)
    try:
        # Find paragraph with '提供を受けた'
        target_para = None
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            if "提供を受けた" in p.Range.Text:
                target_para = p
                break
        if target_para is None:
            return {"error": "target paragraph not found"}

        rng_start = target_para.Range.Start
        rng_end = target_para.Range.End

        # Collect per-char x/y for first ~20 chars
        chars = []
        for i in range(rng_start, min(rng_end, rng_start + 20)):
            r = doc.Range(i, i)
            x = r.Information(wdHorizontal)
            y = r.Information(wdVertical)
            ch = doc.Range(i, i + 1).Text
            chars.append({"i": i - rng_start, "x": x, "y": y, "ch": ch})

        # Per-char advance map (same-line only)
        char_advances = {}
        for j in range(1, len(chars)):
            if chars[j]["y"] == chars[j - 1]["y"]:
                prev_ch = chars[j - 1]["ch"]
                adv = chars[j]["x"] - chars[j - 1]["x"]
                char_advances.setdefault(prev_ch, []).append(adv)

        # Inner kanji set for '提供を受けた匿名': all of '提','供','を','受','け','た','匿','名'
        kanji_set = set('提供を受けた匿名')
        kanji_advs = []
        for ch, advs in char_advances.items():
            if ch in kanji_set:
                kanji_advs.extend(advs)

        # Lines summary
        lines = {}
        for c in chars:
            lines.setdefault(c["y"], []).append(c)
        line_breakdowns = []
        for y_key, lchars in sorted(lines.items()):
            txt = "".join(c["ch"] for c in lchars).rstrip("\r\n").rstrip("\x07")
            line_breakdowns.append({
                "y": y_key,
                "x_start": lchars[0]["x"],
                "x_end": lchars[-1]["x"],
                "text": txt,
                "n_chars": len(lchars),
            })

        # Extract specific advances
        digit_adv = char_advances.get('１', [None])[0]
        dot_adv = char_advances.get('．', [None])[0]
        comma_adv = char_advances.get('，', [None])[0]
        kuten_adv = char_advances.get('。', [None])[0]
        kanji_mean = sum(kanji_advs) / len(kanji_advs) if kanji_advs else None

        return {
            "digit_advance_pt": digit_adv,
            "dot_advance_pt": dot_adv,
            "comma_advance_pt": comma_adv,
            "kuten_advance_pt": kuten_adv,
            "kanji_mean_advance_pt": kanji_mean,
            "kanji_n": len(kanji_advs),
            "lines": line_breakdowns,
        }
    finally:
        doc.Close(SaveChanges=False)


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = {}
    paths = sorted(glob.glob(os.path.join(VARIANTS_DIR, "*.docx")))
    print(f"Measuring {len(paths)} variants...")
    try:
        for i, path in enumerate(paths):
            name = os.path.splitext(os.path.basename(path))[0]
            meta = parse_name(name)
            print(f"  [{i+1}/{len(paths)}] {name}", end="", flush=True)
            try:
                m = measure_doc(word, path)
            except Exception as e:
                print(f"  ERROR: {e}")
                results[name] = {"meta": meta, "error": str(e)}
                continue
            results[name] = {"meta": meta, **m}
            # Show key punct advance
            ch = meta["punct"] if meta else None
            adv = None
            if ch == '．':
                adv = m.get("dot_advance_pt")
            elif ch == '，':
                adv = m.get("comma_advance_pt")
            elif ch == '。':
                adv = m.get("kuten_advance_pt")
            adv_s = f"{adv:.3f}" if adv is not None else "-"
            k_s = f"{m.get('kanji_mean_advance_pt'):.3f}" if m.get("kanji_mean_advance_pt") else "-"
            print(f"  {ch}={adv_s}  k_mean={k_s}")
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT_JSON}")


if __name__ == "__main__":
    main()
