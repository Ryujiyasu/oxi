"""Deep dive: pin the 9pt discrepancy between COM Information(5) and PNG pixel for 1ec1 □.

Memo claims "Word visual x=39pt" (likely COM Information(5)).
User says "PNG pixel x=48pt".

Get ALL geometric data:
  - Shape (textbox) absolute position via Shape.Left
  - Shape size, line position (LineFormat)
  - Shape inset (TextFrame.MarginLeft)
  - First paragraph's first character via Information(5) AND Information(1)
  - Render shape to PNG, find leftmost dark pixel of □
  - Cross-tab to identify which reference frame each is in
"""
import json
import sys
import time
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT = Path("pipeline_data/1ec1_box3_deepdive.json")
PNG = Path("pipeline_data/1ec1_box3_deepdive.png")


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

        # Find Shape 4 (or whichever has □)
        n_shapes = doc.Shapes.Count
        result["n_shapes"] = n_shapes
        result["shapes"] = []
        for si in range(1, n_shapes + 1):
            try:
                shape = doc.Shapes(si)
                shape_name = shape.Name
                if not shape.TextFrame.HasText:
                    continue
                text = (shape.TextFrame.TextRange.Text or "")[:30]
                if "□" not in text:
                    continue

                # Get all geometric properties
                tf = shape.TextFrame
                shape_data = {
                    "shape_idx": si,
                    "shape_name": shape_name,
                    "shape_type": shape.Type,  # msoShapeType enum
                    "text_preview": text,
                    "Left_pt": shape.Left,
                    "Top_pt": shape.Top,
                    "Width_pt": shape.Width,
                    "Height_pt": shape.Height,
                    "MarginLeft_pt": tf.MarginLeft,
                    "MarginRight_pt": tf.MarginRight,
                    "MarginTop_pt": tf.MarginTop,
                    "MarginBottom_pt": tf.MarginBottom,
                    "Orientation": tf.Orientation,
                    "AutoSize": tf.AutoSize,
                    "WordWrap": tf.WordWrap,
                }

                # LineFormat (border)
                try:
                    lf = shape.Line
                    shape_data["Line_Visible"] = lf.Visible
                    shape_data["Line_Weight"] = lf.Weight
                except Exception as e:
                    shape_data["Line_err"] = str(e)

                # Find □ char's COM position
                tr = tf.TextRange
                for ci in range(1, tr.Characters.Count + 1):
                    ch = tr.Characters(ci)
                    if ch.Text == "□":
                        x_info5 = ch.Information(5)  # horizontal pos (typically rel to margin)
                        y_info6 = ch.Information(6)
                        # Try Information(1) (vertical bookkeeping pos)
                        try:
                            info_1 = ch.Information(1)
                        except Exception:
                            info_1 = None
                        try:
                            # MoveStart/Move return char count; Bold/Font for diagnostic
                            font_name = ch.Font.Name
                        except Exception:
                            font_name = None
                        shape_data["box_char"] = {
                            "Information_5_x": x_info5,
                            "Information_6_y": y_info6,
                            "Information_1": info_1,
                            "Font_Name": font_name,
                        }
                        break

                # Try Selection.GoTo + Selection.Information for x
                try:
                    sel = doc.Application.Selection
                    sel.SetRange(tr.Start, tr.Start + 1)
                    shape_data["selection_info5"] = sel.Information(5)
                    shape_data["selection_info6"] = sel.Information(6)
                except Exception as e:
                    shape_data["selection_err"] = str(e)

                result["shapes"].append(shape_data)
                print(f"\nShape {si} ({shape_name}) — Type={shape.Type}")
                print(f"  Left={shape.Left:.2f}pt, Top={shape.Top:.2f}pt, W={shape.Width:.2f}pt, H={shape.Height:.2f}pt")
                print(f"  MarginLeft={tf.MarginLeft:.2f}pt (lIns)")
                if "box_char" in shape_data:
                    print(f"  □ Information(5) = {shape_data['box_char']['Information_5_x']:.2f}pt")
                    print(f"  □ Information(6) = {shape_data['box_char']['Information_6_y']:.2f}pt")
            except Exception as e:
                print(f"Shape {si}: ERR {e}")

        # Render to PNG via Document.SaveAs2 with PDF then... actually use Word's Print to PDF and convert
        # Simpler: get the page that contains the textbox
        print(f"\n--- Saved geometric data ---")

        # Try CopyAsPicture for the shape to clipboard, then save
        # (limited API — typically use SaveAs2 with PDF)

        doc.Close(False)
    finally:
        try: word.Quit()
        except: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
