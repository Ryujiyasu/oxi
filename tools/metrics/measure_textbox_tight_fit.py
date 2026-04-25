"""Measure if Word renders text in tight-fit textboxes.

For each TF_*.docx, open in Word and check the visible text via COM.
Report whether the textbox's text is rendered (visible) in Word.
"""
import win32com.client
import os
import glob

REPRO_DIR = os.path.abspath("tools/metrics/textbox_tight_fit_repro")


def measure_one(word_app, docx_path: str):
    doc = word_app.Documents.Open(docx_path, ReadOnly=True)
    try:
        # Word stores textbox text via Document.Shapes.TextFrame.TextRange.Text
        n_shapes = doc.Shapes.Count
        results = []
        for i in range(1, n_shapes + 1):
            shape = doc.Shapes(i)
            try:
                if shape.TextFrame.HasText:
                    txt = shape.TextFrame.TextRange.Text
                    h = shape.Height  # pt
                    results.append((i, h, txt))
            except:
                pass
        return results
    finally:
        doc.Close(False)


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        for f in sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx'))):
            label = os.path.splitext(os.path.basename(f))[0]
            shapes = measure_one(word, f)
            print(f"\n=== {label} ===")
            for (idx, h, txt) in shapes:
                print(f"  Shape {idx}: height={h:.2f}pt text={txt!r}")
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
