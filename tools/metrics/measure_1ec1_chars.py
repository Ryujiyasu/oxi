"""Measure individual character positions in 1ec1 Shape 4 textbox.

Paragraph-level Range.Information(5) returns -1 for off-page positioned shapes.
But individual character Information(5) may give actual rendered position.
"""
import os, sys, time, json
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT = os.path.abspath("pipeline_data/1ec1_chars_2026-05-02.json")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False; word.DisplayAlerts = False
out = {}
try:
    d = word.Documents.Open(DOC, ReadOnly=True)
    time.sleep(0.5)
    out["page_setup"] = {
        "page_w": float(d.Sections(1).PageSetup.PageWidth),
        "page_margin_l": float(d.Sections(1).PageSetup.LeftMargin),
    }
    n_shapes = d.Shapes.Count
    out["shapes"] = []
    for i in range(1, n_shapes + 1):
        s = d.Shapes(i)
        try:
            tf = s.TextFrame
            if not (tf and tf.HasText):
                continue
            tr = tf.TextRange
            chars = tr.Characters
            char_data = []
            for ci in range(1, min(chars.Count, 30) + 1):
                try:
                    c = chars(ci)
                    char_data.append({
                        "i": ci, "ch": c.Text,
                        "x": float(c.Information(5)),
                        "y": float(c.Information(6)),
                        "page": int(c.Information(3)),
                        "size": float(c.Font.Size),
                    })
                except Exception:
                    pass
            entry = {
                "shape_idx": i,
                "shape_name": s.Name,
                "left": float(s.Left), "top": float(s.Top),
                "width": float(s.Width), "height": float(s.Height),
                "tf_marginL": float(tf.MarginLeft),
                "tf_marginR": float(tf.MarginRight),
                "char_xs": char_data,
            }
            out["shapes"].append(entry)
            print(f"\n[shape {i}] {s.Name}: L={s.Left:.2f} W={s.Width:.2f} "
                  f"tfL={tf.MarginLeft:.2f}")
            for c in char_data[:5]:
                print(f"  ch[{c['i']}] {c['ch']!r} x={c['x']:.2f} y={c['y']:.2f}")
        except Exception as e:
            print(f"[shape {i}] err: {e}")
    d.Close(SaveChanges=False)
finally:
    try: word.Quit()
    except: pass

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2, default=str)
print(f"\nWrote {OUT}")
