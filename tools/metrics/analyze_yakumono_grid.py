"""Session 113 — analyze the yakumono_grid results.

Goal: derive a formula (or model) for '．'/'，'/'。' advance given
(font, fs, cs, line_count, line_chars).

Key questions:
  1. Are '．' '，' '。' always equal? (already observed YES in stdout)
  2. Is the value font-independent at same (fs, cs, line state)?
  3. Does it depend on the line state (n_chars on line, slack)?
  4. Can we predict it from a closed-form formula?
"""
import os
import sys
import io
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
IN_JSON = os.path.normpath(os.path.join(REPO, "tools/metrics/yakumono_grid/results.json"))

with open(IN_JSON, "r", encoding="utf-8") as f:
    data = json.load(f)


def punct_adv(rec):
    p = rec["meta"]["punct"]
    if p == '．':
        return rec.get("dot_advance_pt")
    if p == '，':
        return rec.get("comma_advance_pt")
    if p == '。':
        return rec.get("kuten_advance_pt")
    return None


print("=" * 100)
print("Per-variant: punct adv, kanji_mean, L1 chars, L1 width-ish, ratio punct/kanji")
print("=" * 100)
print(f"{'variant':<42} {'fs':>4} {'cs':>4} {'p':2} {'punct':>6} {'kanji':>6} {'L1n':>4} {'L1tx':<14} {'p/k':>6}")
for name in sorted(data.keys()):
    rec = data[name]
    if "error" in rec or "meta" not in rec:
        continue
    m = rec["meta"]
    p_adv = punct_adv(rec)
    k_adv = rec.get("kanji_mean_advance_pt")
    lines = rec.get("lines", [])
    l1_n = lines[0]["n_chars"] if lines else 0
    l1_tx = (lines[0]["text"] if lines else "")[:12]
    ratio = (p_adv / k_adv) if (p_adv and k_adv) else None
    p_s = f"{p_adv:.2f}" if p_adv else "-"
    k_s = f"{k_adv:.2f}" if k_adv else "-"
    r_s = f"{ratio:.3f}" if ratio else "-"
    print(f"{name:<42} {m['fs_pt']:>4.1f} {m['cs_tw']:>4} {m['punct']} {p_s:>6} {k_s:>6} {l1_n:>4} {l1_tx:<14} {r_s:>6}")

# Cross-font check: at (fs=10.5, cs=-9, dotF), is the value font-independent?
print("\n" + "=" * 100)
print("Cross-font check: fs=10.5, cs=-9, punct='．'")
print("=" * 100)
for slug in ["msmincho", "msgothic", "meiryo", "yugothic", "yumincho"]:
    name = f"g_{slug}_sz21_cs-9_dotF"
    rec = data.get(name)
    if rec and "dot_advance_pt" in rec:
        print(f"  {slug:>10}: '．' = {rec['dot_advance_pt']:.3f}  kanji_mean = {rec.get('kanji_mean_advance_pt'):.3f}")

# Mincho cs-sweep at each fs: detect non-monotonicity
print("\n" + "=" * 100)
print("MS Mincho cs sweep per fs (each row = one fs)")
print("=" * 100)
print(f"{'fs':>5} | {'cs=-5':>8} {'cs=-9':>8} {'cs=-15':>8} {'cs=-20':>8}")
for sz in [18, 21, 24, 28]:
    row = []
    for cs in [-5, -9, -15, -20]:
        name = f"g_msmincho_sz{sz}_cs{cs}_dotF"
        rec = data.get(name)
        if rec:
            v = rec.get("dot_advance_pt")
            row.append(f"{v:.3f}" if v else "-")
        else:
            row.append("-")
    print(f"{sz/2.0:>5.1f} | {row[0]:>8} {row[1]:>8} {row[2]:>8} {row[3]:>8}")

# Same for kanji_mean
print("\nMS Mincho kanji_mean per (fs, cs):")
print(f"{'fs':>5} | {'cs=-5':>8} {'cs=-9':>8} {'cs=-15':>8} {'cs=-20':>8}")
for sz in [18, 21, 24, 28]:
    row = []
    for cs in [-5, -9, -15, -20]:
        name = f"g_msmincho_sz{sz}_cs{cs}_dotF"
        rec = data.get(name)
        if rec:
            v = rec.get("kanji_mean_advance_pt")
            row.append(f"{v:.3f}" if v else "-")
        else:
            row.append("-")
    print(f"{sz/2.0:>5.1f} | {row[0]:>8} {row[1]:>8} {row[2]:>8} {row[3]:>8}")

# Line count per (fs, cs) at MS Mincho dotF
print("\nMS Mincho L1 char count per (fs, cs):")
print(f"{'fs':>5} | {'cs=-5':>8} {'cs=-9':>8} {'cs=-15':>8} {'cs=-20':>8}")
for sz in [18, 21, 24, 28]:
    row = []
    for cs in [-5, -9, -15, -20]:
        name = f"g_msmincho_sz{sz}_cs{cs}_dotF"
        rec = data.get(name)
        if rec:
            lines = rec.get("lines", [])
            v = lines[0]["n_chars"] if lines else 0
            row.append(f"{v}")
        else:
            row.append("-")
    print(f"{sz/2.0:>5.1f} | {row[0]:>8} {row[1]:>8} {row[2]:>8} {row[3]:>8}")

# Line breakdowns
print("\n" + "=" * 100)
print("Line breakdowns per variant (MS Mincho dotF only):")
print("=" * 100)
for sz in [18, 21, 24, 28]:
    for cs in [-5, -9, -15, -20]:
        name = f"g_msmincho_sz{sz}_cs{cs}_dotF"
        rec = data.get(name)
        if not rec or "lines" not in rec:
            continue
        print(f"  fs={sz/2.0:>4.1f} cs={cs}:")
        for ln in rec["lines"][:3]:
            print(f"    L({ln['x_start']:>6.1f}..{ln['x_end']:>6.1f}, {ln['n_chars']:>2} chars): {ln['text']!r}")
