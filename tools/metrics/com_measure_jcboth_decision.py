"""Session 116 — COM-measure punctuation compression decision per
(cell_dxa, text_len). Decision rule: when does Word pick compression?

Output: tools/metrics/jcboth_decision_grid/results.json
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
VARIANTS_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/jcboth_decision_grid/variants"))
OUT_JSON = os.path.normpath(os.path.join(REPO, "tools/metrics/jcboth_decision_grid/results.json"))

wdHorizontal = 5
wdVertical = 6


def parse_name(name):
    m = re.match(r"dg_cw(\d+)_tl(\d+)", name)
    if not m:
        return None
    return {"cell_dxa": int(m.group(1)), "text_len": int(m.group(2))}


def measure_doc(word, path):
    doc = word.Documents.Open(path, ReadOnly=True)
    try:
        target_para = None
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            if "．" in p.Range.Text or "提" in p.Range.Text:
                target_para = p
                break
        if target_para is None:
            return {"error": "target paragraph not found"}

        rng_start = target_para.Range.Start
        rng_end = target_para.Range.End

        chars = []
        for i in range(rng_start, min(rng_end, rng_start + 30)):
            r = doc.Range(i, i)
            x = r.Information(wdHorizontal)
            y = r.Information(wdVertical)
            ch = doc.Range(i, i + 1).Text
            chars.append({"i": i - rng_start, "x": x, "y": y, "ch": ch})

        char_advances = {}
        for j in range(1, len(chars)):
            if chars[j]["y"] == chars[j - 1]["y"]:
                prev_ch = chars[j - 1]["ch"]
                adv = chars[j]["x"] - chars[j - 1]["x"]
                char_advances.setdefault(prev_ch, []).append(adv)

        dot_adv = char_advances.get("．", [None])[0]
        digit_adv = char_advances.get("１", [None])[0]

        # All kanji advances
        kanji_advs = []
        for c, advs in char_advances.items():
            if c not in ("．", "１", "\r", "\x07") and len(c) == 1:
                kanji_advs.extend(advs)
        kanji_mean = sum(kanji_advs) / len(kanji_advs) if kanji_advs else None

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

        return {
            "dot_advance_pt": dot_adv,
            "digit_advance_pt": digit_adv,
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
            d = m.get("dot_advance_pt")
            k = m.get("kanji_mean_advance_pt")
            l1 = m["lines"][0]["n_chars"] if m.get("lines") else 0
            d_s = f"{d:.2f}" if d else "-"
            k_s = f"{k:.2f}" if k else "-"
            print(f"  ．={d_s} k={k_s} L1={l1}")
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT_JSON}")


if __name__ == "__main__":
    main()
