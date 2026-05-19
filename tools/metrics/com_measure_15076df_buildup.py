"""Session 112 — COM-measure '.' advance for each build-up variant.

Iterates over tools/metrics/15076df_buildup/variants/*.docx and reports
the advance of the FULLWIDTH FULL STOP (U+FF0E) at position 1
(content = '１．提供を受けた匿名データの名称').

If a variant shows '．' advance ≈ 6.0pt (vs minimal baseline 9.75pt),
that variant's added content is the trigger.

Output: JSON to tools/metrics/15076df_buildup/results.json
"""
import os
import sys
import io
import json
import glob
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
VARIANTS_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/15076df_buildup/variants"))
OUT_JSON = os.path.normpath(os.path.join(REPO, "tools/metrics/15076df_buildup/results.json"))

wdHorizontal = 5
wdVertical = 6

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
results = {}

try:
    for path in sorted(glob.glob(os.path.join(VARIANTS_DIR, "*.docx"))):
        name = os.path.splitext(os.path.basename(path))[0]
        print(f"\n=== {name} ===")
        try:
            doc = word.Documents.Open(path, ReadOnly=True)
        except Exception as e:
            print(f"  OPEN FAILED: {e}")
            results[name] = {"error": f"open failed: {e}"}
            continue

        try:
            # Find the paragraph with '提供を受けた'
            target_para = None
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                if "提供を受けた" in p.Range.Text:
                    target_para = p
                    break
            if target_para is None:
                print("  TARGET PARA NOT FOUND")
                results[name] = {"error": "target paragraph not found"}
                continue

            rng_start = target_para.Range.Start
            rng_end = target_para.Range.End

            chars = []
            for i in range(rng_start, min(rng_end, rng_start + 20)):
                r = doc.Range(i, i)
                x = r.Information(wdHorizontal)
                y = r.Information(wdVertical)
                ch = doc.Range(i, i + 1).Text
                chars.append({"i": i - rng_start, "x": x, "y": y, "ch": ch})

            # Compute advance of '．' = x[2] - x[1] when both on same line
            char_advances = {}
            for j in range(1, len(chars)):
                if chars[j]["y"] == chars[j - 1]["y"]:
                    prev_ch = chars[j - 1]["ch"]
                    adv = chars[j]["x"] - chars[j - 1]["x"]
                    if prev_ch not in char_advances:
                        char_advances[prev_ch] = []
                    char_advances[prev_ch].append(adv)

            # Specifically '．'
            dot_adv = None
            digit_adv = None
            for c, advs in char_advances.items():
                if c == "．":
                    dot_adv = advs[0]
                if c == "１":
                    digit_adv = advs[0]

            # Detect line breaks
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

            out = {
                "dot_advance_pt": dot_adv,
                "digit_advance_pt": digit_adv,
                "char_advances": {c: advs[0] for c, advs in char_advances.items() if advs},
                "lines": line_breakdowns,
            }
            print(f"  '．' advance: {dot_adv}")
            print(f"  '１' advance: {digit_adv}")
            print(f"  lines: {len(line_breakdowns)}")
            for ln in line_breakdowns:
                print(f"    L: {ln['n_chars']:2d} chars: {ln['text']!r}")
            results[name] = out
        finally:
            doc.Close(SaveChanges=False)
finally:
    word.Quit()

os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2, ensure_ascii=False)

print(f"\nWrote {OUT_JSON}")

# Summary table
print("\n=== Summary ===")
print(f"{'variant':<20} {'.adv':>8} {'1.adv':>8} {'L1 chars':>10}")
for name, r in results.items():
    if "error" in r:
        print(f"{name:<20} ERROR: {r['error']}")
        continue
    da = r.get("dot_advance_pt")
    ia = r.get("digit_advance_pt")
    l1 = r["lines"][0]["n_chars"] if r.get("lines") else "-"
    da_str = f"{da:.3f}" if da is not None else "-"
    ia_str = f"{ia:.3f}" if ia is not None else "-"
    print(f"{name:<20} {da_str:>8} {ia_str:>8} {l1:>10}")
