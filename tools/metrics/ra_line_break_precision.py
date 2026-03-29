"""
Ra: 行折り返し位置の精度をCOM計測で確定
- GDI幅の累積がセル/カラム幅を超える正確なポイント
- Word が行を折り返す判定基準（>=幅? >幅?）
- 文字幅の累積計算とGDI GetTextExtentPoint32W の関係
- 複数run混在行の折り返し判定
"""
import win32com.client, json, os, ctypes

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def gdi_string_width(font_name, ppem, text):
    """Measure full string width (not sum of individual chars)."""
    hdc = user32.GetDC(0)
    hfont = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, font_name)
    old = gdi32.SelectObject(hdc, hfont)
    sz = SIZE()
    gdi32.GetTextExtentPoint32W(hdc, text, len(text), ctypes.byref(sz))
    gdi32.SelectObject(hdc, old)
    gdi32.DeleteObject(hfont)
    user32.ReleaseDC(0, hdc)
    return sz.cx


def gdi_char_widths(font_name, ppem, text):
    """Measure each character individually."""
    hdc = user32.GetDC(0)
    hfont = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, font_name)
    old = gdi32.SelectObject(hdc, hfont)
    widths = []
    for ch in text:
        sz = SIZE()
        gdi32.GetTextExtentPoint32W(hdc, ch, 1, ctypes.byref(sz))
        widths.append(sz.cx)
    gdi32.SelectObject(hdc, old)
    gdi32.DeleteObject(hfont)
    user32.ReleaseDC(0, hdc)
    return widths


def test_linebreak_boundary():
    """Find exact character count where line wraps in a known-width container."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        # Content width = 595.3 - 72 - 72 = 451.3pt = ~601.7px

        # Remove grid
        sectPr_xml = sec.Range.Sections(1).Range.XML
        # Just use no-grid approach by setting large line pitch
        wdoc.Content.Text = ""

        data = {"scenario": "linebreak_boundary", "tests": []}

        for font, fs in [("Calibri", 11), ("Calibri", 9), ("Arial", 11), ("MS Gothic", 10.5)]:
            ppem = round(fs * 96.0 / 72.0)
            content_width_pt = 595.3 - 72 - 72  # A4 page

            # Generate progressively longer text
            base_text = "A" * 200
            char_w_px = gdi_char_widths(font, ppem, "A")[0]
            content_width_px = round(content_width_pt * 96.0 / 72.0)

            # Estimate chars per line
            est_chars = content_width_px // char_w_px

            # Test with exact number of chars
            for n in range(est_chars - 3, est_chars + 4):
                text = "A" * n
                rng = wdoc.Range(0, wdoc.Content.End - 1)
                rng.Delete()
                wdoc.Content.InsertAfter(text)
                para = wdoc.Paragraphs(1)
                para.Range.Font.Name = font
                para.Range.Font.Size = fs
                para.Format.SpaceBefore = 0
                para.Format.SpaceAfter = 0
                wdoc.Repaginate()

                nlines = para.Range.ComputeStatistics(1)  # wdStatisticLines
                string_w = gdi_string_width(font, ppem, text)
                sum_w = sum(gdi_char_widths(font, ppem, text))

                data["tests"].append({
                    "font": font, "size": fs, "ppem": ppem,
                    "n_chars": n, "lines": nlines,
                    "string_width_px": string_w,
                    "sum_char_widths_px": sum_w,
                    "content_width_px": content_width_px,
                })

                if nlines > 1:
                    # Found the break point
                    break

        return data
    finally:
        wdoc.Close(False)


def test_linebreak_mixed_text():
    """Line break with mixed Latin/CJK text."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72
        sec.PageSetup.RightMargin = 72

        data = {"scenario": "linebreak_mixed", "tests": []}

        # Text that should wrap at specific points
        test_texts = [
            "The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog.",
            "Word layout specification test. " * 5,
        ]

        for text in test_texts:
            rng = wdoc.Range(0, wdoc.Content.End - 1)
            rng.Delete()
            wdoc.Content.InsertAfter(text)
            para = wdoc.Paragraphs(1)
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
            wdoc.Repaginate()

            nlines = para.Range.ComputeStatistics(1)

            # Find where each line starts
            prng = para.Range
            lines_info = []
            prev_y = None
            line_start = None
            line_chars = ""

            for ci in range(prng.Start, min(prng.End, prng.Start + 200)):
                cr = wdoc.Range(ci, ci + 1)
                ch = cr.Text
                y = cr.Information(6)
                x = cr.Information(5)

                if prev_y is None or abs(y - prev_y) > 1:
                    if line_chars:
                        lines_info.append({
                            "start_x": round(line_start, 2),
                            "y": round(prev_y, 2),
                            "text": line_chars,
                            "char_count": len(line_chars),
                        })
                    line_start = x
                    line_chars = ""
                if ord(ch) not in (13, 7):
                    line_chars += ch
                prev_y = y

            if line_chars:
                lines_info.append({
                    "start_x": round(line_start, 2),
                    "y": round(prev_y, 2),
                    "text": line_chars,
                    "char_count": len(line_chars),
                })

            # GDI width of each line
            ppem = round(11 * 96.0 / 72.0)
            for li in lines_info:
                li["gdi_string_width_px"] = gdi_string_width("Calibri", ppem, li["text"])
                li["gdi_string_width_pt"] = round(li["gdi_string_width_px"] * 72.0 / 96.0, 2)

            data["tests"].append({
                "text_len": len(text),
                "word_lines": nlines,
                "lines": lines_info,
            })

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_linebreak_boundary()
    results.append(d1)
    print("=== linebreak_boundary ===")
    prev_font = ""
    for t in d1["tests"]:
        if t["font"] != prev_font:
            print(f"\n  {t['font']} {t['size']}pt (ppem={t['ppem']}):")
            print(f"  content_width = {t['content_width_px']}px")
            prev_font = t["font"]
        marker = " <-- WRAPS" if t["lines"] > 1 else ""
        print(f"    n={t['n_chars']}: lines={t['lines']}, "
              f"string_w={t['string_width_px']}px, sum_chars={t['sum_char_widths_px']}px{marker}")

    d2 = test_linebreak_mixed_text()
    results.append(d2)
    print(f"\n=== linebreak_mixed ===")
    for ti, t in enumerate(d2["tests"]):
        print(f"\n  Text {ti+1} ({t['text_len']} chars, {t['word_lines']} lines):")
        content_w_pt = 595.3 - 72 - 72
        for li in t["lines"]:
            over = "OVER" if li["gdi_string_width_pt"] > content_w_pt else "ok"
            print(f"    y={li['y']}: {li['char_count']}ch, "
                  f"gdi_w={li['gdi_string_width_pt']}pt/{content_w_pt:.1f}pt [{over}]")
            print(f"      \"{li['text'][:60]}...\"" if len(li['text']) > 60 else f"      \"{li['text']}\"")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_line_break_precision.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
