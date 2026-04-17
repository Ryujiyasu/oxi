"""
Per-character x-measurement for b837 paras[48] in Word vs Oxi.

Gives a char-by-char layout of where Word places each character.
Compared with Oxi dump, identifies the exact point where wrap decisions diverge.
"""
import io, json, time, sys
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path(r"C:/Users/ryuji/oxi-4/tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
OUT = Path(__file__).with_name("output") / "b837_para48_charx.json"
OUT.parent.mkdir(parents=True, exist_ok=True)


def main():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False

    try:
        doc = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
        time.sleep(1.5)

        target_idx = 49  # 1-based
        p = doc.Paragraphs(target_idx)
        rng = p.Range

        print(f"Paragraph {target_idx}: chars = {rng.End - rng.Start}")

        sel = word.Selection
        chars_data = []
        for ci in range(rng.Start, rng.End):
            sel.SetRange(ci, ci + 1)
            try:
                y = sel.Information(6)
                x = sel.Information(5)
                pg = int(sel.Information(3))
                ch = sel.Text
            except Exception:
                continue
            chars_data.append({
                "ci": ci,
                "page": pg,
                "y": round(y, 2),
                "x": round(x, 2),
                "ch": ch,
                "codepoint": ord(ch) if len(ch) == 1 else None,
            })

        doc.Close(False)

        # Group by (page, y_key)
        from collections import defaultdict
        lines = defaultdict(list)
        for c in chars_data:
            y_key = round(c['y'] * 2) / 2
            lines[(c['page'], y_key)].append(c)

        print(f"\n{len(chars_data)} chars total, {len(lines)} lines:")
        for key in sorted(lines.keys()):
            chars = lines[key]
            chars.sort(key=lambda c: c['x'])
            pg, y = key
            text = ''.join(c['ch'] for c in chars)
            # Compute widths (next char's x - this char's x)
            widths = []
            for i in range(len(chars) - 1):
                widths.append(round(chars[i+1]['x'] - chars[i]['x'], 2))
            widths.append(None)  # last char no next
            print(f"\n  page={pg} y={y:.1f} chars={len(chars)}")
            for i, c in enumerate(chars):
                w = f"w={widths[i]}" if widths[i] is not None else "w=(last)"
                print(f"    [{i:2d}] x={c['x']:6.2f} {w:>9s}  ch={c['ch']!r} (U+{c['codepoint']:04X})" if c['codepoint'] else
                      f"    [{i:2d}] x={c['x']:6.2f} {w:>9s}  ch={c['ch']!r}")

        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(chars_data, f, indent=2, ensure_ascii=False)
        print(f"\nSaved -> {OUT}")
    finally:
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
