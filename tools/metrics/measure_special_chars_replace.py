"""Take special_chars_spacing_01.docx, replace text with various probes,
keeping all other settings/styles intact. See if the 11.5pt step pattern
reproduces with simple text."""
import win32com.client
import os
import sys
import tempfile
import zipfile
import re

sys.stdout.reconfigure(encoding="utf-8")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

ORIG = os.path.abspath("pipeline_data/docx/special_chars_spacing_01.docx")


def replace_text_in_docx(new_text):
    tmp = os.path.join(tempfile.gettempdir(), "special_chars_replaced.docx")
    if os.path.exists(tmp):
        os.remove(tmp)
    with zipfile.ZipFile(ORIG, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "word/document.xml":
                    s = data.decode("utf-8")
                    # Replace contents of <w:t> with new_text
                    s = re.sub(
                        r'(<w:t[^>]*>)([^<]*)(</w:t>)',
                        lambda m: m.group(1) + new_text + m.group(3),
                        s, count=1,
                    )
                    data = s.encode("utf-8")
                zout.writestr(item, data)
    return tmp


def measure(new_text):
    tmp = replace_text_in_docx(new_text)
    doc = word.Documents.Open(tmp, ReadOnly=True)
    chars = doc.Range().Characters
    out = []
    prev_x = None
    prev_line = None
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            ln = c.Information(10)
            x = c.Information(5)
            dx = (x - prev_x) if (prev_x is not None and ln == prev_line) else None
            out.append((ch, ln, x, dx))
            prev_x = x
            prev_line = ln
        except Exception:
            pass
    doc.Close(SaveChanges=False)
    try:
        os.remove(tmp)
    except Exception:
        pass
    return out


def report(label, text):
    data = measure(text)
    line1 = [(ch, x, dx) for ch, ln, x, dx in data if ln == 1]
    from collections import Counter
    hist = Counter(round(dx, 2) for _, _, dx in line1 if dx is not None)
    last_x = line1[-1][1] if line1 else 0
    print(f"{label:40s}  L1={len(line1):2d}  start→last={last_x:6.2f}  dx={dict(hist)}")
    if hist and len(hist) > 1:
        # Show outliers
        modal = hist.most_common(1)[0][0]
        odd = [(i, ch, dx) for i, (ch, _, dx) in enumerate(line1)
               if dx is not None and abs(dx - modal) > 0.01]
        for i, ch, dx in odd[:8]:
            ctx = line1[max(0, i-1)][0]
            print(f"    [{i:2d}] {ctx!r}→{ch!r}  dx={dx:.2f}")


# Tests with text replacement preserving all original styles
report("漢×50 (in original doc)", "漢" * 50)
report("ア×50", "ア" * 50)
report("a×50 (latin)", "a" * 50)
report("特殊文字：×8 + 漢×30", "特殊文字：" * 8 + "漢" * 10)
# Very simple alternating
report("漢ア漢ア×25", "漢ア" * 25)
report("ABab×12 + 漢×26", "ABab" * 6 + "漢" * 26)
# Mixed punctuation
report("漢、漢×25", "漢、" * 25)
report("漢。漢×25", "漢。" * 25)
# Pure kanji + 1 touten at end
report("漢×40 + 、",            "漢" * 40 + "、")
report("漢×40 + 。",            "漢" * 40 + "。")
# Just kanji+touten (NO 3rd char)
report("漢、 ×25",              "漢、" * 25)
# Touten only
report("、×50",                  "、" * 50)
# Test with punctuation density
report("漢漢漢、 ×13",          "漢漢漢、" * 13)
report("漢漢漢漢漢、×9",         "漢漢漢漢漢、" * 9)
# Test: pure kanji with multi-line para (does Word stretch L1 to fit 40?)
report("漢×80 (2-line para)",   "漢" * 80)
report("漢×120 (3-line para)",  "漢" * 120)
# vs single-line short para
report("漢×30 (single line)",   "漢" * 30)
# Toggle: 39 vs 40 boundary with explicit
report("漢×39 + ×0",            "漢" * 39)
report("漢×40 + ×0 (natural)",  "漢" * 40)

# The original
report("ORIGINAL special_chars text",
       "特殊文字：①②③④⑤⑥⑦⑧⑨⑩　記号：★☆◆◇■□●○　単位：㎡㎏㎝　括弧：【】『』〈〉《》")

word.Quit()
