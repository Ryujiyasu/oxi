"""Bisect Word's stretch+hang trigger condition by varying CJK punct density."""
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


def replace_text(new_text):
    tmp = os.path.join(tempfile.gettempdir(), "bisect_stretch.docx")
    if os.path.exists(tmp):
        os.remove(tmp)
    with zipfile.ZipFile(ORIG, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "word/document.xml":
                    s = data.decode("utf-8")
                    s = re.sub(
                        r'(<w:t[^>]*>)([^<]*)(</w:t>)',
                        lambda m: m.group(1) + new_text + m.group(3),
                        s, count=1,
                    )
                    data = s.encode("utf-8")
                zout.writestr(item, data)
    return tmp


def measure(new_text):
    tmp = replace_text(new_text)
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
    try: os.remove(tmp)
    except: pass
    return out


def report(label, text):
    data = measure(text)
    line1 = [(ch, x, dx) for ch, ln, x, dx in data if ln == 1]
    from collections import Counter
    hist = Counter(round(dx, 2) for _, _, dx in line1 if dx is not None)
    last_x = line1[-1][1] if line1 else 0
    n_punct = sum(1 for ch, _, _ in line1 if ch in "、。，．：；？！「」『』（）〔〕【】《》〈〉")
    density = n_punct / len(line1) * 100 if line1 else 0
    is_stretch = "STRETCH" if 11.5 in hist else "       "
    print(f"{label:30s} L1={len(line1):2d} pos40={last_x:6.2f} punct={n_punct:2d}/{len(line1):2d}={density:5.1f}% {is_stretch}  dx={dict(hist)}")


# Bisect punct density
print("=== Density bisection ===")
report("漢×50 (0%)",          "漢" * 50)
report("漢×9+、 ×5 (10%)",     ("漢" * 9 + "、") * 5)        # 1/10
report("漢×7+、 ×6 (12.5%)",   ("漢" * 7 + "、") * 6)        # 1/8
report("漢×6+、 ×7 (14.3%)",   ("漢" * 6 + "、") * 7)        # 1/7
report("漢×5+、 ×8 (16.7%)",   ("漢" * 5 + "、") * 8)        # 1/6
report("漢×4+、 ×10 (20%)",    ("漢" * 4 + "、") * 10)       # 1/5
report("漢×3+、 ×12 (25%)",    ("漢" * 3 + "、") * 12)       # 1/4
report("漢×2+、 ×16 (33%)",    ("漢" * 2 + "、") * 16)       # 1/3
report("漢×1+、 ×25 (50%)",    ("漢" * 1 + "、") * 25)       # 1/2

print("\n=== Position-only tests ===")
# Same density 25% (1 punct per 4 chars), different positions
report("dense at start",  "、" * 10 + "漢" * 30)
report("dense at end",    "漢" * 30 + "、" * 10)
report("scattered 25%",   ("漢漢漢、") * 10)
report("alternating",     ("漢、") * 20)

print("\n=== ）」 hangable check ===")
report("）density 25%",    ("漢漢漢）") * 10)
report("」density 25%",    ("漢漢漢」") * 10)
report("。density 25%",    ("漢漢漢。") * 10)
report(",density 25% (ASCII)", ("漢漢漢,") * 10)

word.Quit()
