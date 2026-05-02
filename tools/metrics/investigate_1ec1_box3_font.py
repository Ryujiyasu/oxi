"""Investigation A: get Word's actual font for 1ec1's □3 paragraph (in textbox).

Per session_50_1ec1_phase3_findings.md, Oxi `□` glyph is 2.4pt LEFT of Word's.
Hypothesis H_A: Oxi resolves theme major eastAsia to wrong font (e.g., Yu
Mincho Light vs Word's actual font).

Steps:
  1. Open 1ec1091177b1_006.docx in Word
  2. Find textbox shape containing □3 (Shape 4 / TB[4])
  3. Get the paragraph's first run Font.Name
  4. Get Font.NameAsian, Font.NameOther for full picture
  5. Also check Range.Font.Name on actual char range
"""
import json
import sys
import time
import zipfile
import re
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT = Path("pipeline_data/1ec1_box3_font_word.json")


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    result = {}
    try:
        doc = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
        time.sleep(0.5)
        # Iterate Shapes
        n_shapes = doc.Shapes.Count
        print(f"Doc has {n_shapes} shapes")
        for si in range(1, n_shapes + 1):
            try:
                shape = doc.Shapes(si)
                shape_name = shape.Name
                # Has textframe?
                if not shape.TextFrame.HasText:
                    continue
                tf = shape.TextFrame
                text_range = tf.TextRange
                full_text = (text_range.Text or "")[:50]
                print(f"\nShape {si} ({shape_name}): text='{full_text!r}'")
                # Find paragraph containing □ and 3
                if "□" not in full_text and "Shape 4" not in shape_name:
                    continue
                # For matching shape, walk paragraphs
                n_p = text_range.Paragraphs.Count
                shape_data = {"shape_idx": si, "shape_name": shape_name, "text": full_text, "paragraphs": []}
                for pi in range(1, n_p + 1):
                    p = text_range.Paragraphs(pi)
                    pr = p.Range
                    p_text = (pr.Text or "")[:50].replace("\r", "\\r").replace("\x07", "\\x07")
                    if "□" not in p_text:
                        continue
                    # Get font of first run with content
                    first_char = pr.Characters(1)
                    font = first_char.Font
                    p_data = {
                        "para_idx": pi,
                        "text": p_text,
                        "Font.Name": font.Name,
                        "Font.NameAscii": getattr(font, "NameAscii", None),
                        "Font.NameOther": getattr(font, "NameOther", None),
                        "Font.NameFarEast": getattr(font, "NameFarEast", None),
                        "Font.NameBi": getattr(font, "NameBi", None),
                        "Font.Size": font.Size,
                        "Font.Bold": font.Bold,
                    }
                    # Also get the □ char specifically
                    for ci in range(1, pr.Characters.Count + 1):
                        ch = pr.Characters(ci)
                        ch_text = ch.Text
                        if ch_text == "□":
                            ch_font = ch.Font
                            p_data["box_char"] = {
                                "Font.Name": ch_font.Name,
                                "Font.NameAscii": getattr(ch_font, "NameAscii", None),
                                "Font.NameOther": getattr(ch_font, "NameOther", None),
                                "Font.NameFarEast": getattr(ch_font, "NameFarEast", None),
                                "Font.Size": ch_font.Size,
                            }
                            break
                    print(f"  Para {pi}: '{p_text}'")
                    print(f"    Font.Name (first char): {font.Name}")
                    print(f"    NameAscii={getattr(font, 'NameAscii', '?')}, NameFarEast={getattr(font, 'NameFarEast', '?')}, NameOther={getattr(font, 'NameOther', '?')}")
                    if "box_char" in p_data:
                        print(f"    □ char Font.Name: {p_data['box_char']['Font.Name']}")
                        print(f"    □ NameFarEast: {p_data['box_char'].get('Font.NameFarEast', '?')}")
                    shape_data["paragraphs"].append(p_data)
                if shape_data["paragraphs"]:
                    result.setdefault("shapes", []).append(shape_data)
            except Exception as e:
                print(f"Shape {si} ERR: {e}")

        # Also extract theme info from docx
        with zipfile.ZipFile(DOCX) as z:
            try:
                theme = z.read("word/theme/theme1.xml").decode("utf-8", errors="replace")
                # Find majorFont eastAsia
                major_match = re.search(r'<a:majorFont>.*?<a:ea\s+typeface="([^"]*)"', theme, re.DOTALL)
                minor_match = re.search(r'<a:minorFont>.*?<a:ea\s+typeface="([^"]*)"', theme, re.DOTALL)
                result["theme"] = {
                    "majorFont.eastAsia": major_match.group(1) if major_match else "(empty)",
                    "minorFont.eastAsia": minor_match.group(1) if minor_match else "(empty)",
                }
                print(f"\nTheme:")
                print(f"  majorFont eastAsia: {result['theme']['majorFont.eastAsia']!r}")
                print(f"  minorFont eastAsia: {result['theme']['minorFont.eastAsia']!r}")
            except Exception as e:
                print(f"Theme extract ERR: {e}")
        doc.Close(False)
    finally:
        try: word.Quit()
        except: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
