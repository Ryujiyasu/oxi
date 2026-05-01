"""Measure SR_* relativeFrom variants."""
import json, re, zipfile, sys
from pathlib import Path
import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\sp_relfrom_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\sp_relfrom_measurements.json")


def parse_relfrom_and_offset(p):
    with zipfile.ZipFile(p) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    m = re.search(r'<wp:positionV\s+relativeFrom="([^"]+)">\s*<wp:posOffset>(-?\d+)</wp:posOffset>', xml, re.S)
    if not m: return None, None
    return m.group(1), int(m.group(2))


def main():
    docs = sorted(REPRO_DIR.glob("SR_*.docx"))
    print(f"Measuring {len(docs)} docs...", file=sys.stderr)
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in docs:
            relfrom, posV_emu = parse_relfrom_and_offset(d)
            posV_pt = posV_emu / 12700.0 if posV_emu is not None else 0
            try:
                doc = word.Documents.Open(str(d.resolve()), ReadOnly=True)
            except Exception as e:
                results.append({"file": d.name, "error": f"open: {e}"})
                continue
            try:
                if doc.Shapes.Count == 0:
                    results.append({"file": d.name, "error": "no shape"})
                    continue
                shape = doc.Shapes(1)
                shape_top = shape.Top
                rel_v = shape.RelativeVerticalPosition
                anchor_top = None
                anchor_text = None
                try:
                    ap = shape.Anchor.Paragraphs(1)
                    anchor_top = ap.Range.Information(6)
                    anchor_text = (ap.Range.Text or "")[:30].replace("\r","\\r")
                except Exception: pass
                # First body para Y for "page" / "margin" reference
                first_y = doc.Paragraphs(1).Range.Information(6)
                results.append({
                    "file": d.name,
                    "xml_relativeFrom": relfrom,
                    "xml_posOffset_emu": posV_emu,
                    "xml_posOffset_pt": posV_pt,
                    "shape_top_offset_pt": shape_top,
                    "rel_v_pos_int": rel_v,
                    "anchor_top": anchor_top,
                    "anchor_text": anchor_text,
                    "first_body_para_y": first_y,
                })
                print(f"  done {d.name}", file=sys.stderr)
            except Exception as e:
                results.append({"file": d.name, "error": f"measure: {e}"})
            finally:
                try: doc.Close(SaveChanges=0)
                except Exception: pass
    finally:
        try: word.Quit()
        except Exception: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)

    print()
    print(f"{'file':28s} {'xml_rel':>15} {'xml_pt':>7} {'rvi':>3} {'s.Top':>7} {'a_top':>7} {'1st_y':>7}")
    print("-" * 95)
    for r in results:
        if 'error' in r:
            print(f"  {r['file'][:26]:26s}  ERROR: {r['error']}")
            continue
        print(f"  {r['file'][:26]:26s}"
              f" {r['xml_relativeFrom']:>15s}"
              f" {r['xml_posOffset_pt']:>7.2f}"
              f" {r['rel_v_pos_int']:>3d}"
              f" {r['shape_top_offset_pt']:>7.2f}"
              f" {(r['anchor_top'] or 0):>7.2f}"
              f" {r['first_body_para_y']:>7.2f}")


if __name__ == "__main__":
    main()
