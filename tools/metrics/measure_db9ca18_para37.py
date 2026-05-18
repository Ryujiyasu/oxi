"""S107: measure db9ca18 paragraph 37 to verify Word's break behavior at +5.25pt overflow."""
import sys
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
DOCX = ROOT / 'tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx'


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(str(DOCX.absolute()), ReadOnly=True)
        try:
            paras = list(doc.Paragraphs)
            print(f"Total paragraphs: {len(paras)}")
            # paragraph 37 (0-based or 1-based?)
            for idx in [35, 36, 37, 38, 39]:
                p = paras[idx]
                rng = p.Range
                start = rng.Start
                end = rng.End
                first_y = doc.Range(start, start).Information(6)
                first_page = doc.Range(start, start).Information(3)
                txt = rng.Text.rstrip('\r\n')[:40]
                print(f"\npi={idx}: pg={first_page} first_y={first_y:.2f} len={len(rng.Text.rstrip())} text={txt!r}")
                # measure each line
                last_y = None
                line_idx = 0
                for off in range(min(end - start, 1000)):
                    char_rng = doc.Range(start + off, start + off + 1)
                    t = char_rng.Text
                    if t in ('\r', '\n'):
                        continue
                    y = char_rng.Information(6)
                    pg = char_rng.Information(3)
                    if last_y is None or abs(y - last_y) > 1.0:
                        print(f"    line {line_idx}: off={off:3d} pg={pg} y={y:7.2f} char={t!r}")
                        line_idx += 1
                    last_y = y
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
