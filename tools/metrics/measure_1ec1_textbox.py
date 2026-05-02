"""1ec1 TextBox COM measurement (Task C).

For 1ec1091177b1_006.docx, enumerate doc.Shapes(*) and capture:
  - Shape.Name / Type
  - Anchor: Left, Top, Width, Height (from page or paragraph)
  - TextFrame: MarginLeft/Top/Right/Bottom (= bodyPr lIns/tIns/rIns/bIns)
  - First paragraph's first char position (Information(5) horizontal)
  - First word text content (to identify which is the body bullet)

Goal: discover the +1.81pt right-shift origin for the `□3` body bullet.
"""
import os
import sys
import json
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC_PATH = os.path.abspath(
    "tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT = os.path.abspath("pipeline_data/1ec1_textbox_com_2026-05-02.json")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out = {"doc": DOC_PATH, "shapes": []}
    try:
        d = word.Documents.Open(DOC_PATH, ReadOnly=True)
        time.sleep(0.5)
        # Section margins
        sec = d.Sections(1).PageSetup
        out["page_setup"] = {
            "page_w": float(sec.PageWidth),
            "page_h": float(sec.PageHeight),
            "margin_left": float(sec.LeftMargin),
            "margin_right": float(sec.RightMargin),
            "margin_top": float(sec.TopMargin),
            "margin_bottom": float(sec.BottomMargin),
        }
        print(f"\nSection 1 page_setup:", json.dumps(out["page_setup"],
                                                       indent=2))

        # Enumerate all shapes
        n_shapes = d.Shapes.Count
        print(f"\nTotal shapes: {n_shapes}")
        for i in range(1, n_shapes + 1):
            s = d.Shapes(i)
            entry = {
                "index": i,
                "name": s.Name,
                "type": int(s.Type),
                "left": float(s.Left),
                "top": float(s.Top),
                "width": float(s.Width),
                "height": float(s.Height),
                "anchor_text": "",
                "has_text_frame": False,
            }
            # Anchor info
            try:
                anchor = s.Anchor
                entry["anchor_text"] = anchor.Text[:50] if anchor and anchor.Text else ""
                entry["anchor_para_idx"] = None
            except Exception as e:
                entry["anchor_err"] = str(e)
            # Text frame margins
            try:
                tf = s.TextFrame
                if tf and tf.HasText:
                    entry["has_text_frame"] = True
                    entry["margin_left"] = float(tf.MarginLeft)
                    entry["margin_top"] = float(tf.MarginTop)
                    entry["margin_right"] = float(tf.MarginRight)
                    entry["margin_bottom"] = float(tf.MarginBottom)
                    # First paragraph in textbox
                    txt_range = tf.TextRange
                    paras = []
                    for pi in range(1, txt_range.Paragraphs.Count + 1):
                        p = txt_range.Paragraphs(pi)
                        prange = p.Range
                        ptxt = prange.Text[:80]
                        # First char x
                        try:
                            x_first = float(prange.Information(5))
                            y_first = float(prange.Information(6))
                        except Exception:
                            x_first = None; y_first = None
                        paras.append({
                            "idx": pi,
                            "text": ptxt,
                            "first_char_x": x_first,
                            "first_char_y": y_first,
                            "left_indent": float(prange.ParagraphFormat.LeftIndent),
                            "first_line_indent": float(prange.ParagraphFormat.FirstLineIndent),
                        })
                    entry["paragraphs"] = paras
                else:
                    entry["has_text_frame"] = bool(tf and tf.HasText)
            except Exception as e:
                entry["text_frame_err"] = str(e)

            out["shapes"].append(entry)
            print(f"\n[shape {i}] name='{s.Name}' type={int(s.Type)} "
                  f"L={entry['left']:.2f} T={entry['top']:.2f} "
                  f"W={entry['width']:.2f} H={entry['height']:.2f}")
            if entry.get("has_text_frame"):
                print(f"   TextFrame margins: L={entry['margin_left']:.2f} "
                      f"T={entry['margin_top']:.2f} R={entry['margin_right']:.2f} "
                      f"B={entry['margin_bottom']:.2f}")
                for p in entry.get("paragraphs", []):
                    print(f"     para[{p['idx']}] x={p['first_char_x']} "
                          f"y={p['first_char_y']} li={p['left_indent']:.2f} "
                          f"fli={p['first_line_indent']:.2f} "
                          f"text={p['text']!r}")
            print(f"   anchor.text={entry['anchor_text']!r}")

        # Also: dump body paragraphs near the relevant area
        print("\n=== Body paragraphs ===")
        body_paras = []
        for pi in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(pi)
            r = p.Range
            txt = r.Text[:60]
            try:
                xa = float(r.Information(5))
                ya = float(r.Information(6))
            except Exception:
                xa = ya = None
            body_paras.append({"idx": pi, "x": xa, "y": ya, "text": txt})
        out["body_paragraphs"] = body_paras
        for bp in body_paras[:30]:
            print(f"  [body p{bp['idx']:>3}] x={bp['x']} y={bp['y']} "
                  f"text={bp['text']!r}")

        d.Close(SaveChanges=False)
    finally:
        try: word.Quit()
        except: pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2, default=str)
    print(f"\nWrote {OUT}", flush=True)


if __name__ == "__main__":
    main()
