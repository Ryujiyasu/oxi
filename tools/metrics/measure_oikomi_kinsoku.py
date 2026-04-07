"""Measure Word's resolution strategy for line-start-prohibited characters.

Question: When a line-start-prohibited char (e.g. ）」、。) would naturally land
at the start of the next line, does Word:
  (A) hang it past the right margin (burasagari, ぶら下げ)
  (B) push the previous char(s) down to the next line (押し出し/oikomi)

Method:
- Generate text where the natural break point lands exactly on a target char.
- US Letter, 90pt L/R margins → content width = 432pt.
- ＭＳ 明朝 10.5pt full-width → 10.5pt per CJK char → 41 chars max per line.
- Place the test char at position 42 to force overflow at that exact char.
- Read Word's per-char line number via Information(10) (wdFirstCharacterLineNumber).
- Report L1 length for each pattern.

Interpretations of L1 length when test char is at pos 42:
  L1 = 42  → Word hung the char (ぶら下げ)
  L1 = 41  → char treated as non-prohibited; standard break-before
  L1 = 40  → 1-char push down (押し出し by 1)
  L1 < 40  → multi-char push down chain
"""
import win32com.client
import time
import sys

sys.stdout.reconfigure(encoding="utf-8")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

WD_FIRST_LINE = 10  # wdFirstCharacterLineNumber


def measure(text, font="ＭＳ 明朝", size=10.5):
    """Return list of (char, line_no) and L1 length (chars on line 1)."""
    doc = word.Documents.Add()
    time.sleep(0.1)
    # US Letter, 90pt L/R margins → 432pt content (matches ruby_text_lineheight_11)
    ps = doc.PageSetup
    ps.PageWidth = 612.0
    ps.PageHeight = 792.0
    ps.LeftMargin = 90.0
    ps.RightMargin = 90.0
    ps.TopMargin = 72.0
    ps.BottomMargin = 72.0
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0  # left
    time.sleep(0.05)
    chars = doc.Range().Characters
    per_char = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            ln = c.Information(WD_FIRST_LINE)
            per_char.append((ch, ln))
        except Exception:
            continue
    doc.Close(SaveChanges=False)

    # Count chars per line
    line_lens = {}
    for _, ln in per_char:
        line_lens[ln] = line_lens.get(ln, 0) + 1
    return per_char, line_lens


# Tests: place trigger char at exactly position 42 (the first overflow position).
# Filler = 41 normal CJK chars (漢) — line max is 41 chars at 10.5pt full-width.
F41 = "漢" * 41

PATTERNS = [
    # ===== Single line-start-prohibited char at pos 42 =====
    ("baseline_42_normal", F41 + "漢"),         # char 42 = 漢 (not prohibited)
    ("kuten_42",          F41 + "。"),         # 。
    ("touten_42",         F41 + "、"),         # 、
    ("close_paren_42",    F41 + "）"),         # ）
    ("close_bracket_42",  F41 + "」"),         # 」
    ("close_bracket2_42", F41 + "』"),         # 』
    ("close_kakko_42",    F41 + "〕"),         # 〕
    ("colon_42",          F41 + "："),         # ：
    ("semi_42",           F41 + "；"),         # ；
    ("excl_42",           F41 + "！"),         # ！
    ("ques_42",           F41 + "？"),         # ？
    ("chouon_42",         F41 + "ー"),         # ー (allowed start in Word default after cb2c27b)
    ("small_ya_42",       F41 + "ゃ"),         # small kana
    # ===== Two prohibited chars at pos 42, 43 =====
    ("close_close_42",    "漢" * 41 + "）」"), # 41 normal + ）」
    ("close_kuten_42",    "漢" * 41 + "）。"), # ）。
    ("kuten_kuten_42",    "漢" * 41 + "。。"),
    # ===== Three prohibited chars =====
    ("triple_close_42",   "漢" * 41 + "）」』"),
    # ===== Line-end-prohibited (opening parens at line end) =====
    # Place ( at pos 41 so the natural break (before pos 42) would put ( at line end.
    ("open_paren_at_41",  "漢" * 40 + "（漢"),  # pos 41 = （, pos 42 = 漢
    ("open_brkt_at_41",   "漢" * 40 + "「漢"),  # pos 41 = 「, pos 42 = 漢
    # ===== Pattern from ruby_text_lineheight_11 (full text, first 50 chars) =====
    ("ruby_lh11_actual",
     "漢字にルビを付けた文章：「専門用語（せんもんようご）」「技術革新（ぎじゅつかくしん）」「情報処理（じょうほうしょり）」が含まれる段落で、行間の自動拡張動作を検証します。"),
]

print(f"{'pattern':25s}  {'L1':>3s}  {'L2':>3s}  {'L3':>3s}  notes")
print("-" * 70)
for name, text in PATTERNS:
    try:
        per_char, line_lens = measure(text)
    except Exception as e:
        print(f"{name:25s}  ERR: {e}")
        continue
    lines_sorted = sorted(line_lens.items())
    L = [v for _, v in lines_sorted]
    L_str = "  ".join(f"{x:3d}" for x in L[:5])
    # Show last char of L1 and first char of L2 if present
    notes = ""
    if len(L) >= 2:
        L1_end_idx = L[0]
        if L1_end_idx <= len(per_char):
            l1_last = per_char[L1_end_idx - 1][0]
            l2_first = per_char[L1_end_idx][0] if L1_end_idx < len(per_char) else "?"
            notes = f"L1.end={l1_last!r} L2.start={l2_first!r}"
    print(f"{name:25s}  {L_str}  {notes}")

word.Quit()
