"""Measure TR_* trHeight matrix via Word COM Tables(1).Rows(1).Height."""
import json, re, sys
from pathlib import Path
import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\tr_height_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\tr_height_measurements.json")


def main():
    docs = sorted(REPRO_DIR.glob("TR_*.docx"))
    print(f"Measuring {len(docs)} docs...", file=sys.stderr)
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in docs:
            try:
                doc = word.Documents.Open(str(d.resolve()), ReadOnly=True)
            except Exception as e:
                results.append({"file": d.name, "error": f"open: {e}"})
                continue
            try:
                tbl = doc.Tables(1)
                row = tbl.Rows(1)
                set_height_pt = row.Height
                rule = row.HeightRule  # 0=Auto, 1=AtLeast, 2=Exactly
                tt = tbl.Range.Information(6)
                tp = tbl.Range.Information(3)
                cell = tbl.Cell(1, 1)
                cell_first_y = cell.Range.Paragraphs(1).Range.Information(6)
                # Post-table paragraph y (rendered row bottom proxy)
                # body has: anchor para + table + tail para
                # tail = Paragraphs after the inline table.
                # Strategy: take the last paragraph in the body and measure its top y;
                # rendered row height ≈ tail_para_top - table_top.
                tail_y = None
                try:
                    tail_para = doc.Paragraphs(doc.Paragraphs.Count)
                    tail_y = tail_para.Range.Information(6)
                except Exception:
                    pass
                rendered = (tail_y - tt) if (tail_y and tt) else None
                results.append({
                    "file": d.name,
                    "row_set_height_pt": set_height_pt,
                    "row_height_rule_int": rule,
                    "row_height_rule": ["auto","atLeast","exact"][rule] if rule in (0,1,2) else "?",
                    "table_top_pt": tt,
                    "table_page": tp,
                    "first_cell_first_para_y": cell_first_y,
                    "tail_para_y": tail_y,
                    "rendered_height_proxy_pt": rendered,
                })
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

    # Pretty-print sorted by rule, lines, lp, tw
    print()
    print(f"{'file':38s} {'rule':>7} {'spec':>5} {'set':>5} {'rend':>5} {'tail_y':>7} {'tbl_y':>7}")
    print("-" * 90)
    for r in sorted(results, key=lambda x: x['file']):
        if 'error' in r:
            print(f"  {r['file']:36s}  ERROR: {r['error']}")
            continue
        m = re.match(r'TR_(\w+?)_h(\d+)_(\d+)L_lp(\d+)', r['file'])
        if not m: continue
        rule, tw, lines, lp = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4))
        spec_pt = tw / 20.0
        rh = r.get('rendered_height_proxy_pt')
        print(f"  {r['file']:36s}"
              f" {r['row_height_rule']:>7s}"
              f" {spec_pt:>5.1f}"
              f" {r['row_set_height_pt']:>5.1f}"
              f" {(f'{rh:5.1f}' if rh is not None else '   - '):>5}"
              f" {(r.get('tail_para_y') or 0):>7.2f}"
              f" {r['table_top_pt']:>7.2f}")

    # Group by rule + lines + lp; print rendered heights per tw
    print()
    print("=== Rendered row height proxy vs trHeight value (spec_pt) ===")
    grouped = {}
    for r in results:
        if 'error' in r: continue
        m = re.match(r'TR_(\w+?)_h(\d+)_(\d+)L_lp(\d+)', r['file'])
        if not m: continue
        rule, tw, lines, lp = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4))
        rh = r.get('rendered_height_proxy_pt')
        grouped.setdefault((rule, lines, lp), []).append((tw / 20.0, rh))
    for k in sorted(grouped):
        rule, lines, lp = k
        pts = sorted(grouped[k])
        print(f"  rule={rule:8s} lines={lines} lp={lp}: {pts}")


if __name__ == "__main__":
    main()
