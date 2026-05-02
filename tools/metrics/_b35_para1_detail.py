"""Verify what b35 para 1 actually contains and what its Information(6)
reports — to resolve the apparent contradiction between §3.3 (Info(6)
= glyph top) and §13.6 verification (Info(6) = line-box top).
"""
import os, sys, time
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import win32com.client

DOCX_REAL = "C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
DOC = "b35123fe8efc_tokumei_08_01.docx"


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)

    path = os.path.join(DOCX_REAL, DOC)
    last_err = None
    for attempt in range(3):
        try:
            wdoc = word.Documents.Open(path)
            break
        except Exception as e:
            last_err = e
            time.sleep(2)
            try:
                while word.Documents.Count > 0:
                    word.Documents(1).Close(False)
            except Exception:
                pass
    else:
        print(last_err); return

    try:
        wdoc.Repaginate()
        time.sleep(0.5)

        # First 5 paragraphs detail
        for i in range(1, 6):
            p = wdoc.Paragraphs(i)
            r = p.Range
            text = r.Text[:50]
            y_para = round(r.Information(6), 4)
            try:
                fc = r.Characters(1)
                y_char = round(fc.Information(6), 4)
            except Exception:
                y_char = None
            font = r.Font
            font_name = font.Name
            font_size = font.Size
            line_spacing = p.Format.LineSpacing
            line_spacing_rule = p.Format.LineSpacingRule
            print(f"para {i}:")
            print(f"  text: {text!r}")
            print(f"  font: {font_name} {font_size}pt")
            print(f"  lineSpacing: {line_spacing} rule={line_spacing_rule}")
            print(f"  para Information(6): {y_para}")
            print(f"  first char Information(6): {y_char}")
            if y_char is not None and abs(y_char - y_para) > 0.01:
                print(f"  diff (char - para): {y_char - y_para:+.2f}pt")
            print()
    finally:
        wdoc.Close(False)
        try: word.Quit()
        except Exception: pass


if __name__ == "__main__":
    main()
