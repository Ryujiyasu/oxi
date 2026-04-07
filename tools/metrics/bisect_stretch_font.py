"""Test if stretch trigger depends on font/size by using ruby_text_lineheight_11
as the template (ＭＳ 明朝 10.5pt) instead of special_chars (メイリオ 11pt)."""
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


def measure_in_template(template, new_text):
    tmp = os.path.join(tempfile.gettempdir(), "bisect_font.docx")
    if os.path.exists(tmp):
        os.remove(tmp)
    with zipfile.ZipFile(template, "r") as zin:
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
    doc = word.Documents.Open(tmp, ReadOnly=True)
    chars = doc.Range().Characters
    out = []
    prev_x = None; prev_line = None
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci); ch = c.Text
            if ch in ("\r","\x07"): continue
            ln = c.Information(10); x = c.Information(5)
            dx = (x - prev_x) if (prev_x is not None and ln == prev_line) else None
            out.append((ch, ln, x, dx))
            prev_x = x; prev_line = ln
        except: pass
    doc.Close(False)
    try: os.remove(tmp)
    except: pass
    return out


def report(template_name, label, text):
    template = os.path.abspath(f"pipeline_data/docx/{template_name}.docx")
    data = measure_in_template(template, text)
    line1 = [(ch, x, dx) for ch, ln, x, dx in data if ln == 1]
    from collections import Counter
    hist = Counter(round(dx, 2) for _, _, dx in line1 if dx is not None)
    last_x = line1[-1][1] if line1 else 0
    is_stretch = "STRETCH" if any(d != 11.0 and d != 10.5 for d in hist if d is not None) and len(hist) > 1 else "       "
    char_at_n = line1[-1][0] if line1 else "?"
    print(f"  {label:30s} L1={len(line1):2d} char[L1]={char_at_n!r} last_x={last_x:6.2f}  {is_stretch}  dx={dict(hist)}")


print("=== Template: ruby_text_lineheight_11 (ＭＳ 明朝 10.5pt) ===")
report("ruby_text_lineheight_11", "漢×50",            "漢" * 50)
report("ruby_text_lineheight_11", "漢×3+、×15",      "漢漢漢、" * 15)
report("ruby_text_lineheight_11", "漢×3+）×15",      "漢漢漢）" * 15)
report("ruby_text_lineheight_11", "漢×3+。×15",      "漢漢漢。" * 15)
report("ruby_text_lineheight_11", "漢×40+）",         "漢" * 40 + "）")
report("ruby_text_lineheight_11", "漢×41+）",         "漢" * 41 + "）")
report("ruby_text_lineheight_11", "漢×40+、",         "漢" * 40 + "、")

print("\n=== Template: special_chars_spacing_01 (メイリオ 11pt) ===")
report("special_chars_spacing_01", "漢×50",           "漢" * 50)
report("special_chars_spacing_01", "漢×3+、×13",     "漢漢漢、" * 13)
report("special_chars_spacing_01", "漢×3+）×13",     "漢漢漢）" * 13)

print("\n=== Para continuation effect (MS Mincho 10.5pt) ===")
# Same first-line content but different para tail
report("ruby_text_lineheight_11", "漢×41+）alone",      "漢" * 41 + "）")
report("ruby_text_lineheight_11", "漢×41+）+漢×60",     "漢" * 41 + "）" + "漢" * 60)
report("ruby_text_lineheight_11", "漢×41+）+漢×3",      "漢" * 41 + "）" + "漢" * 3)
# What about ）at exactly position 41 instead of 42
report("ruby_text_lineheight_11", "漢×40+）+漢×60",     "漢" * 40 + "）" + "漢" * 60)
# The actual ruby_lh11 first 42 chars
report("ruby_text_lineheight_11", "ruby first 50",
       "漢字にルビを付けた文章：「専門用語（せんもんようご）」「技術革新（ぎじゅつかくしん）し")
report("ruby_text_lineheight_11", "ruby first 50 + tail",
       "漢字にルビを付けた文章：「専門用語（せんもんようご）」「技術革新（ぎじゅつかくしん）し" + "漢" * 60)
report("ruby_text_lineheight_11", "ruby FULL ORIGINAL",
       "漢字にルビを付けた文章：「専門用語（せんもんようご）」「技術革新（ぎじゅつかくしん）」「情報処理（じょうほうしょり）」が含まれる段落で、行間の自動拡張動作を検証します。次の行との間隔についても確認が必要です。")

print("\n=== Tail length effect (when L1 has stretch candidate at char 42) ===")
# 漢×41+）+漢×N — does L2 length change strategy?
PREFIX = "漢" * 41 + "）"  # naturally creates char 42 = ）
for tail_n in [0, 1, 2, 3, 5, 10, 20, 41, 60]:
    report("ruby_text_lineheight_11",
           f"漢×41+）+漢×{tail_n}",
           PREFIX + "漢" * tail_n)

print("\n=== Effect of char 43 type (multi-bracket sequence) ===")
report("ruby_text_lineheight_11", "漢×41+）+漢",         "漢" * 41 + "）漢")
report("ruby_text_lineheight_11", "漢×41+）+」",         "漢" * 41 + "）」")
report("ruby_text_lineheight_11", "漢×41+）+」+漢×60",   "漢" * 41 + "）」" + "漢" * 60)
report("ruby_text_lineheight_11", "漢×41+）+。+漢×60",   "漢" * 41 + "）。" + "漢" * 60)
report("ruby_text_lineheight_11", "漢×41+）+、+漢×60",   "漢" * 41 + "）、" + "漢" * 60)
report("ruby_text_lineheight_11", "漢×41+）+「+漢×60",   "漢" * 41 + "）「" + "漢" * 60)
# What if 2 closing brackets in a row at the natural break
report("ruby_text_lineheight_11", "漢×40+）」+漢×60",    "漢" * 40 + "）」" + "漢" * 60)
report("ruby_text_lineheight_11", "漢×40+）」",           "漢" * 40 + "）」")

print("\n=== Reproduce chain test results in ruby_lh11 template ===")
# These had L1=40 (oikomi) in chain test (created new doc + roundtrip)
# Check if same text in ruby_lh11 template gives different result
report("ruby_text_lineheight_11", "F41+）漢漢漢",   "漢" * 41 + "）漢漢漢")
report("ruby_text_lineheight_11", "F41+。漢漢漢",   "漢" * 41 + "。漢漢漢")
report("ruby_text_lineheight_11", "F41+、漢漢漢",   "漢" * 41 + "、漢漢漢")
report("ruby_text_lineheight_11", "F41+」漢漢漢",   "漢" * 41 + "」漢漢漢")
report("ruby_text_lineheight_11", "F41+）」漢漢",  "漢" * 41 + "）」漢漢")
report("ruby_text_lineheight_11", "F41+。。漢漢",  "漢" * 41 + "。。漢漢")

word.Quit()
