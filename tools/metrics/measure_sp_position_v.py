"""Measure SP_* shape positionV repros via Word COM.

For each repro, find the wp:anchor's containing paragraph (the "anchor
paragraph"), measure its top y, then measure Shape.Top (Word COM
property in points).

Predicted: shape_top = anchor_top + posOffset_pt.
"""
import json, re, zipfile, sys
from pathlib import Path
import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\sp_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\sp_position_v_measurements.json")


def parse_posOffset_emu(p):
    with zipfile.ZipFile(p) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    m = re.search(r'<wp:positionV[^>]*>\s*<wp:posOffset>(-?\d+)</wp:posOffset>', xml, re.S)
    return int(m.group(1)) if m else None


def main():
    docs = sorted(REPRO_DIR.glob("SP_*.docx"))
    print(f"Measuring {len(docs)} docs...", file=sys.stderr)
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in docs:
            posV_emu = parse_posOffset_emu(d)
            posV_pt  = posV_emu / 12700.0 if posV_emu is not None else None
            try:
                doc = word.Documents.Open(str(d.resolve()), ReadOnly=True)
            except Exception as e:
                results.append({"file": d.name, "error": f"open: {e}"})
                continue
            try:
                # Locate the shape (Word's first floating shape)
                shapes = doc.Shapes
                if shapes.Count == 0:
                    results.append({"file": d.name, "error": "no shapes"})
                    continue
                shape = shapes(1)
                shape_top = shape.Top  # offset relative to RelativeVerticalPosition reference
                shape_left = shape.Left
                rel_v_pos = shape.RelativeVerticalPosition
                rel_h_pos = shape.RelativeHorizontalPosition
                # Find anchor paragraph: shape.Anchor is a Range; use it
                anchor_top = None
                anchor_text = None
                try:
                    anchor_range = shape.Anchor
                    anchor_para = anchor_range.Paragraphs(1)
                    anchor_top = anchor_para.Range.Information(6)
                    anchor_text = (anchor_para.Range.Text or "")[:30].replace("\r","\\r").replace("\x07","\\x07")
                except Exception as e:
                    pass
                # All paragraphs preceding the anchor for context
                all_paras = []
                for i in range(1, doc.Paragraphs.Count + 1):
                    p = doc.Paragraphs(i)
                    txt = (p.Range.Text or "")[:30].replace("\r","\\r").replace("\x07","\\x07")
                    all_paras.append({"i": i, "y": p.Range.Information(6), "text": txt})
                # shape.Top is OFFSET from rel_v_pos reference (paragraph top),
                # so absolute Y = anchor_top + shape.Top
                absolute_top = (anchor_top + shape_top) if anchor_top is not None else None
                results.append({
                    "file": d.name,
                    "posV_emu": posV_emu,
                    "posV_pt": posV_pt,
                    "shape_top_offset_pt": shape_top,
                    "shape_left_pt": shape_left,
                    "rel_v_pos": rel_v_pos,
                    "rel_h_pos": rel_h_pos,
                    "anchor_para_top_pt": anchor_top,
                    "anchor_para_text": anchor_text,
                    "absolute_top_pt": absolute_top,
                    "shape_top_minus_posV": (shape_top - posV_pt) if posV_pt is not None else None,
                    "all_paras": all_paras,
                })
                print(f"  done {d.name}", file=sys.stderr)
            except Exception as e:
                results.append({"file": d.name, "error": f"measure: {e}"})
                print(f"  ERROR {d.name}: {e}", file=sys.stderr)
            finally:
                try: doc.Close(SaveChanges=0)
                except Exception: pass
    finally:
        try: word.Quit()
        except Exception: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print()
    print(f"{'file':25s} {'posV':>7} {'a_top':>7} {'s.Top':>7} {'abs':>7} {'resid':>7} {'rvpos':>5}")
    print("-" * 80)
    for r in results:
        if 'error' in r:
            print(f"  {r['file'][:23]:23s}  ERROR: {r['error']}")
            continue
        resid = r.get('shape_top_minus_posV', 0) or 0
        print(f"  {r['file'][:23]:23s}"
              f" {r['posV_pt']:>7.2f}"
              f" {(r['anchor_para_top_pt'] or 0):>7.2f}"
              f" {r['shape_top_offset_pt']:>7.2f}"
              f" {(r['absolute_top_pt'] or 0):>7.2f}"
              f" {resid:>+7.2f}"
              f" {r['rel_v_pos']:>5d}")


if __name__ == "__main__":
    main()
