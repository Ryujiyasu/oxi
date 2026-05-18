"""S107: measure each character's y position in d77a pi=25 to find line break.

Goal: determine exactly where Word breaks pi=25 between pages.
Also: get Word's reported y for specific lines (LBT or glyph_top).
"""
import sys
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
DOCX = ROOT / 'tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx'


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(str(DOCX.absolute()), ReadOnly=True)
        try:
            target = None
            for i, p in enumerate(doc.Paragraphs):
                txt = p.Range.Text
                if '従来のように政府標準利用規約' in txt:
                    target = p
                    print(f"pi=?, idx={i}, len={len(txt)}, first40={txt[:40]}")
                    break
            if target is None:
                print("Not found!"); return
            rng = target.Range
            start = rng.Start
            end = rng.End
            print(f"\nMeasuring each character position...")
            last_y = None
            line_idx = 0
            line_first_chars = []
            for off in range(end - start):
                char_rng = doc.Range(start + off, start + off + 1)
                txt = char_rng.Text
                if txt in ('\r', '\n'):
                    continue
                y = char_rng.Information(6)
                page = char_rng.Information(3)
                if last_y is None or abs(y - last_y) > 1.0:
                    line_first_chars.append((line_idx, off, txt, page, y))
                    line_idx += 1
                last_y = y
            print(f"  Total lines detected: {len(line_first_chars)}")
            for li, off, txt, page, y in line_first_chars:
                print(f"  line {li}: offset {off:3d} pg={page} y={y:7.2f} char={txt!r}")
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
