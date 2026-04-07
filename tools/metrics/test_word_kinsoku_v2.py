"""V2: Does Word do word-aware kinsoku?

Test: insert a real word and see if Word breaks INSIDE the word at small kana,
or KEEPS the word together.
"""
import win32com.client
import time
import sys

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def test_lines(text, font="ＭＳ 明朝", size=10.5):
    doc = word.Documents.Add()
    time.sleep(0.2)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.1)
    chars = doc.Range().Characters
    prev_y = None
    line_chars = []
    out_lines = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            cy = c.Information(6)
            if prev_y is not None and abs(cy - prev_y) > 0.5:
                out_lines.append("".join(line_chars))
                line_chars = []
            prev_y = cy
            line_chars.append(ch)
        except:
            continue
    if line_chars:
        out_lines.append("".join(line_chars))
    doc.Close(SaveChanges=False)
    return out_lines


# Cases:
# 1. Pure repetition: positions where break happens
tests = [
    # (label, text)
    ("rep_a_yo_i", "あ" * 40 + "ョ" + "い" * 20),     # artificial
    ("rep_a_komuni", "あ" * 35 + "コミュニケーション" + "あ" * 5),  # real word
    ("rep_a_my", "あ" * 40 + "ュ" + "い" * 20),         # ュ alone
    ("rep_komuni_inline", "システム開発における要件定義フェーズでは、ステークホルダーとの密接なコミュニケーションが極めて重要"),  # natural sentence
    ("rep_a_okurigana", "あ" * 40 + "っ" + "た" + "あ" * 20),  # っ + た (okurigana)
]

for label, text in tests:
    lines = test_lines(text)
    sys.stdout.buffer.write(f"\n=== {label} ===\n".encode("utf-8", "replace"))
    for i, l in enumerate(lines):
        sys.stdout.buffer.write(f"  L{i} ({len(l)}ch): {l}\n".encode("utf-8", "replace"))

word.Quit()
