"""Investigate the 0.5pt step pattern observed in special_chars_spacing_01.

Hypothesis tests:
  1. Is it font-specific (Meiryo vs other)?
  2. Is it position-based (cumulative rounding)?
  3. Is it char-pair-specific (kerning)?
  4. Does it depend on font_size?
"""
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


def make_docx(text, font="メイリオ", sz_halfpt=22):
    """Build a docx with given font/size, US Letter, 90pt L/R margins."""
    doc = word.Documents.Add()
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
    rng.Font.Size = sz_halfpt / 2.0
    doc.Paragraphs(1).Alignment = 0
    tmp = os.path.join(tempfile.gettempdir(), "meiryo_test.docx")
    if os.path.exists(tmp):
        os.remove(tmp)
    doc.SaveAs2(tmp, FileFormat=12)
    doc.Close(SaveChanges=False)
    return tmp


def measure(text, font="メイリオ", sz_halfpt=22):
    """Build, save with doNotCompress, reopen, return list of (ch, line, x, dx)."""
    tmp = make_docx(text, font, sz_halfpt)
    # Inject doNotCompress to disable yakumono compression
    tmp2 = tmp + ".edited.docx"
    with zipfile.ZipFile(tmp, "r") as zin:
        with zipfile.ZipFile(tmp2, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "word/settings.xml":
                    s = data.decode("utf-8")
                    if "characterSpacingControl" in s:
                        s = re.sub(
                            r'<w:characterSpacingControl[^/]*/>',
                            '<w:characterSpacingControl w:val="doNotCompress"/>',
                            s,
                        )
                    else:
                        s = s.replace(
                            "</w:settings>",
                            '<w:characterSpacingControl w:val="doNotCompress"/></w:settings>',
                        )
                    data = s.encode("utf-8")
                zout.writestr(item, data)
    doc = word.Documents.Open(tmp2, ReadOnly=True)
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
        os.remove(tmp2)
    except Exception:
        pass
    return out


def report(label, text, font="メイリオ", sz_halfpt=22):
    print(f"\n=== {label}  font={font} sz={sz_halfpt/2}pt ===")
    data = measure(text, font, sz_halfpt)
    # Show only line 1 (ignore wraps)
    line1 = [(ch, x, dx) for ch, ln, x, dx in data if ln == 1]
    # Build dx histogram
    from collections import Counter
    hist = Counter(round(dx, 2) for _, _, dx in line1 if dx is not None)
    print(f"  L1 chars: {len(line1)}  width: {line1[-1][1] - line1[0][1] if line1 else 0:.2f}pt (start→last char)")
    print(f"  dx histogram: {dict(hist)}")
    # Show chars that have non-modal dx
    if hist:
        modal = hist.most_common(1)[0][0]
        odd = [(i, ch, dx) for i, (ch, _, dx) in enumerate(line1) if dx is not None and abs(dx - modal) > 0.01]
        if odd:
            print(f"  modal={modal}  outliers ({len(odd)}):")
            for i, ch, dx in odd[:10]:
                ctx = line1[max(0,i-1)][0] if i>0 else "?"
                print(f"    [{i}] {ctx!r}→{ch!r}  dx={dx:.2f}")


# Tests
report("kanji×50 meiryo 11pt", "漢" * 50, "メイリオ", 22)
report("kanji×50 ms_mincho 11pt", "漢" * 50, "ＭＳ 明朝", 22)
report("kanji×50 ms_gothic 11pt", "漢" * 50, "ＭＳ ゴシック", 22)
report("kanji×50 yu_mincho 11pt", "游明朝", 22)
report("kanji×50 meiryo 10.5pt", "漢" * 50, "メイリオ", 21)
report("kanji×50 ms_mincho 10.5pt", "漢" * 50, "ＭＳ 明朝", 21)

# The exact special_chars_spacing_01 text
report("special_chars meiryo 11pt",
       "特殊文字：①②③④⑤⑥⑦⑧⑨⑩　記号：★☆◆◇■□●○　単位：㎡㎏㎝　括弧：【】『』〈〉《》",
       "メイリオ", 22)

word.Quit()
