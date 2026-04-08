"""Sweep all CJK punctuation to determine full yakumono compression set.

Tests each candidate char in two contexts:
  A. Single between CJK ideographs (漢X字) → expect full-width 11.0
  B. Doubled with itself (XX) → if 2nd advance < 11.0, char compresses
  C. Adjacent to (（) → cross-pair compression

Output: per-char {single, doubled, paired_with_paren} advances.
"""
import win32com.client
import time
import json
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

# Candidate yakumono characters (CJK punctuation block + brackets)
CANDIDATES = [
    # Brackets
    "（", "）", "「", "」", "『", "』", "【", "】",
    "〔", "〕", "｛", "｝", "〈", "〉", "《", "》",
    "［", "］",
    # Punctuation
    "、", "。", "，", "．", "・", "：", "；", "！", "？",
    # Quotes
    "“", "”", "‘", "’",
    # Long-vowel / dashes
    "ー", "—", "―",
    # Middle dot
    "／", "＼",
]

def char_advances(text, font="ＭＳ 明朝", size=11.0):
    doc = word.Documents.Add()
    time.sleep(0.15)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.05)
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            xs.append((ch, c.Information(5)))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    advs = []
    for i in range(len(xs) - 1):
        advs.append((xs[i][0], round(xs[i+1][1] - xs[i][1], 4)))
    return advs

results = {}
for ch in CANDIDATES:
    try:
        a_single = char_advances(f"漢{ch}字{ch}漢")  # isolated and repeated-far
        a_double = char_advances(f"漢{ch}{ch}漢")    # adjacent pair
        a_triple = char_advances(f"{ch}{ch}{ch}{ch}")  # pure run of 4
    except Exception as e:
        results[ch] = {"error": str(e)}
        continue
    results[ch] = {
        "isolated":  a_single,
        "doubled":   a_double,
        "run_of_4":  a_triple,
    }
    # Brief one-line summary
    run = a_triple
    advs = [w for _, w in run]
    print(f"{ch} U+{ord(ch):04X}: run4={advs}")

word.Quit()

out_path = "tools/metrics/output/yakumono_sweep.json"
import os
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f"\nSaved: {out_path}")
