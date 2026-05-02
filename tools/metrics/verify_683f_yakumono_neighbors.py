"""Verify R17 hypothesis: 683f p2 L1 has 、 followed by CJK ideograph
(which would mean FINAL RULE B→CJK predicts no compression — and R17's
list_marker gate is a workaround, not the real rule).
"""
import win32com.client
import time
import sys
import os

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(
    "tools/golden-test/documents/docx/"
    "683ffcab86e2_20230331_resources_open_data_contract_addon_00.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
try:
    d = word.Documents.Open(DOC, ReadOnly=True)
    time.sleep(0.4)
    # Find paragraph 2 (R17 cited "p2 p3")
    for pi in [2, 3]:
        try:
            p = d.Paragraphs(pi)
            text = p.Range.Text
            print(f"\n=== P{pi} text (first 80 chars) ===")
            print(repr(text[:80]))
            print(f"  len={len(text)}")
            chars = p.Range.Characters
            xs = []
            for ci in range(1, min(chars.Count + 1, 80)):
                try:
                    c = chars(ci)
                    ch = c.Text
                    if ch in ("\r", "\x07"):
                        continue
                    xs.append((ci, ch, float(c.Information(5)),
                               float(c.Information(6)),
                               c.Font.Name, c.Font.Size))
                except Exception:
                    continue
            # Print first line only
            if not xs:
                continue
            y0 = xs[0][3]
            print(f"  L1 chars (y={y0}):")
            for ci, ch, x, y, fn, sz in xs:
                if abs(y - y0) > 0.5:
                    break
                marker = ""
                if ch in "、。」』）】〕｝〉》］，．":
                    marker = "  <-- yakumono Type B"
                if ch in "「『（【〔｛〈《［":
                    marker = "  <-- yakumono Type A"
                print(f"    [{ci:2d}] {ch!r:>5} x={x:7.2f} font={fn} sz={sz}"
                      + marker)
            # advance per char on L1
            l1 = [r for r in xs if abs(r[3] - y0) < 0.5]
            print("  Advances (L1):")
            for i in range(len(l1) - 1):
                adv = round(l1[i + 1][2] - l1[i][2], 4)
                ch = l1[i][1]
                next_ch = l1[i + 1][1]
                marker = ""
                if ch in "、。」』）】〕｝〉》］，．":
                    marker = f"  <-- B chr, neighbor='{next_ch}'"
                print(f"    pos{l1[i][0]} {ch!r}->{next_ch!r} adv={adv}"
                      + marker)
        except Exception as e:
            print(f"P{pi} error: {e}")
    d.Close(SaveChanges=False)
finally:
    word.Quit()
