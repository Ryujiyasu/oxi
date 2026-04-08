"""Measure autoSpaceDE boundary spacing across font sizes.

Hypothesis: extra = fontSize / 4 (one-quarter em).
At 10pt → 2.5, 10.5pt → 2.625, 11pt → 2.75, 12pt → 3.0, 14pt → 3.5, 18pt → 4.5
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(font_cjk, font_latin, size):
    doc = word.Documents.Add()
    time.sleep(0.15)
    rng = doc.Range()
    rng.InsertAfter("はM")  # CJK kana then Latin letter
    rng = doc.Range()
    rng.Font.Size = size
    # Use single multi-name list (Word picks per-script automatically)
    rng.Font.Name = font_cjk
    time.sleep(0.05)
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r","\x07"):
                continue
            xs.append((ch, c.Information(5), c.Font.Name))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    if len(xs) >= 2:
        adv = xs[1][1] - xs[0][1]
        return adv, xs
    return None, xs

def measure_text(text, font, size):
    doc = word.Documents.Add()
    time.sleep(0.15)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    time.sleep(0.05)
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r","\x07"):
                continue
            xs.append((ch, c.Information(5), c.Font.Name))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    return xs

print("CJK→Latin boundary spacing (はM) per CJK font:")
print(f"{'cjk_font':<14} {'size':<6} {'CJK_adv':<8} {'extra':<8}")
CJK_FONTS = ["ＭＳ 明朝", "ＭＳ ゴシック", "Yu Mincho", "Yu Gothic", "Meiryo"]
SIZES = [9.0, 10.0, 10.5, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0]
for cjk in CJK_FONTS:
    for size in SIZES:
        try:
            adv, xs = measure(cjk, "Times New Roman", size)
            extra = adv - size
            print(f"{cjk:<14} {size:<6} {adv:<8.3f} {extra:+.3f}")
        except Exception as e:
            print(f"{cjk:<14} {size:<6} ERR {e}")
    print()

word.Quit()
sys.exit(0)
print("\nLatin→CJK boundary spacing (Mは):")
print("size  M_adv (Times New Roman M expected)  extra(vs natural)")
for size in [9.0, 10.0, 10.5, 11.0, 12.0, 14.0, 16.0, 18.0]:
    xs = measure_text("Mは", "ＭＳ 明朝", size)
    if len(xs) >= 2:
        m_adv = xs[1][1] - xs[0][1]
        # Get M's natural width via "MM" baseline
        xs_mm = measure_text("MM", "ＭＳ 明朝", size)
        m_natural = xs_mm[1][1] - xs_mm[0][1] if len(xs_mm) >= 2 else 0
        extra = m_adv - m_natural
        print(f"{size:>5}  M_adv={m_adv:>6.3f}  M_natural={m_natural:>6.3f}  extra={extra:>+5.3f}")

print("\nLatin word boundary (test は):")
print("size  t_adv  extra(vs MM_adv-style natural)")
for size in [10.0, 11.0, 12.0]:
    xs = measure_text("testは", "ＭＳ 明朝", size)
    print(f"  size={size}:")
    for i in range(len(xs)-1):
        adv = xs[i+1][1] - xs[i][1]
        print(f"    {xs[i][0]!r}: adv={adv:.3f}")

word.Quit()
