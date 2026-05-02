# -*- coding: utf-8 -*-
"""Inspect Shape 9 TextFrame margins via Word COM API."""
import sys, os, time
import pythoncom, win32com.client as wc
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")

pythoncom.CoInitialize()
word = None
for attempt in range(5):
    try:
        word = wc.Dispatch("Word.Application")
        time.sleep(2)
        word.Visible = False
        word.DisplayAlerts = False
        break
    except Exception as e:
        print(f"Word startup {attempt+1}: {e}")
        time.sleep(6)
if word is None:
    print("Failed Word"); sys.exit(1)

try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    print(f"Opened: {doc.Name}")

    for i in range(1, doc.Shapes.Count + 1):
        shape = doc.Shapes(i)
        try: name = shape.Name
        except: name = "?"
        try:
            tf = shape.TextFrame
            text = tf.TextRange.Text[:50] if tf.HasText else ""
        except: text = ""
        if '□' in (text or ''):
            print(f"\n=== Shape {i}: {name!r} ===")
            print(f"  Width={shape.Width} Height={shape.Height}")
            print(f"  Left={shape.Left} Top={shape.Top}")
            try:
                print(f"  HorizontalPositionRelativeTo={shape.RelativeHorizontalPosition}")
            except: pass
            try: print(f"  WrapFormat.Type={shape.WrapFormat.Type}")
            except: pass
            try: print(f"  TextFrame.MarginLeft={tf.MarginLeft}")
            except Exception as e: print(f"  MarginLeft err: {e}")
            try: print(f"  TextFrame.MarginTop={tf.MarginTop}")
            except: pass
            try: print(f"  TextFrame.MarginRight={tf.MarginRight}")
            except: pass
            try: print(f"  TextFrame.MarginBottom={tf.MarginBottom}")
            except: pass
            try: print(f"  TextFrame.WordWrap={tf.WordWrap}")
            except: pass
            try: print(f"  TextFrame.AutoSize={tf.AutoSize}")
            except: pass
            try: print(f"  TextFrame2.MarginLeft={shape.TextFrame2.MarginLeft}")
            except Exception as e: print(f"  TextFrame2 err: {e}")
            try: print(f"  TextFrame2.HorizontalAnchor={shape.TextFrame2.HorizontalAnchor}")
            except: pass
            try: print(f"  TextFrame2.VerticalAnchor={shape.TextFrame2.VerticalAnchor}")
            except: pass
            try: print(f"  TextFrame2.WordWrap={shape.TextFrame2.WordWrap}")
            except: pass
            try: print(f"  TextFrame2.AutoSize={shape.TextFrame2.AutoSize}")
            except: pass
            # Inspect all paragraphs
            try:
                p_count = tf.TextRange.Paragraphs.Count
                print(f"  Paragraphs: {p_count}")
                for pi in range(1, min(p_count + 1, 4)):
                    p = tf.TextRange.Paragraphs(pi)
                    pf = p.Range.ParagraphFormat
                    txt = p.Range.Text[:30]
                    print(f"    P{pi}: text={txt!r}")
                    print(f"        LeftIndent={pf.LeftIndent} CharacterUnitLeftIndent={getattr(pf, 'CharacterUnitLeftIndent', '?')}")
                    print(f"        FirstLineIndent={pf.FirstLineIndent}")
            except Exception as e: print(f"  Para err: {e}")
    doc.Close(SaveChanges=False)
finally:
    try: word.Quit()
    except: pass
