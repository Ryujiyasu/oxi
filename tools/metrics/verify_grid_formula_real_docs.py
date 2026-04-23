"""Verify grid line-height formula against real baseline docs.

For each doc, iterate first N paragraphs, collect (font_size, lineRule, Y, gap_to_next),
then compute predicted gap via formula and flag matches.

Formula: gap = (ceil(N/grid)*grid + N)/2, N = round(fs * 1.2)
"""
import win32com.client, time, math, json
from pathlib import Path

DOCS = [
    ("d77a58485f16_20240705_resources_data_outline_08.docx", 18.0),
    ("9a8e8ddab85b_order_06-1.docx",                         18.0),
    ("15076df085f5_tokumei_08_09.docx",                      16.8),
    ("bd90b00ab7a7_order_05.docx",                           16.5),
]
BASE = Path("tools/golden-test/documents/docx")
OUT = Path(__file__).parent / "grid_formula_real_docs.json"


def predict_gap(fs_pt, grid_pt):
    N = math.floor(fs_pt * 1.2 + 0.5)
    eff = math.ceil(N / grid_pt) * grid_pt
    return (eff + N) / 2.0, N, eff


def measure(word, path: Path, grid_pt: float, max_paras: int = 30):
    doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
    time.sleep(0.5)
    rows = []
    try:
        ycache = {}
        for i in range(1, min(doc.Paragraphs.Count, max_paras) + 1):
            try:
                p = doc.Paragraphs(i).Range
                y = doc.Range(p.Start, p.Start + 1).Information(6)
                # font size (use first char's font)
                fs = p.Characters(1).Font.Size if p.Characters.Count > 0 else None
                # line rule
                try:
                    rule = p.ParagraphFormat.LineSpacingRule
                except Exception:
                    rule = None
                # line spacing value
                try:
                    ls = p.ParagraphFormat.LineSpacing
                except Exception:
                    ls = None
                # snap to grid
                try:
                    snap = p.ParagraphFormat.DisableLineHeightGrid
                except Exception:
                    snap = None
                # Number of runs / text preview
                text = p.Text[:25] if p.Text else ""
                rows.append(dict(idx=i, y=round(y,3), fs=fs, rule=rule, ls=ls, snap=snap, text=text))
            except Exception as e:
                rows.append(dict(idx=i, error=str(e)))
    finally:
        doc.Close(False)

    # compute gaps
    for i in range(len(rows) - 1):
        if "y" in rows[i] and "y" in rows[i+1]:
            rows[i]["gap_to_next"] = round(rows[i+1]["y"] - rows[i]["y"], 3)
    return rows


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    results = {}
    try:
        for fname, grid_pt in DOCS:
            path = BASE / fname
            if not path.exists():
                print(f"MISSING: {fname}")
                continue
            print(f"\n=== {fname}  grid={grid_pt}pt ===")
            rows = measure(word, path, grid_pt)
            # report: for each row, predicted gap per font size
            for r in rows:
                if "error" in r or "gap_to_next" not in r:
                    continue
                fs = r.get("fs")
                if fs is None or fs <= 0: continue
                pred_gap, N, eff = predict_gap(fs, grid_pt)
                match = "OK" if abs(pred_gap - r["gap_to_next"]) < 0.05 else "  "
                print(f"  p{r['idx']:>3d}  fs={fs:>4.1f}  gap={r['gap_to_next']:>6.2f}  pred={pred_gap:>5.2f}  N={N:>2d} eff={eff:>4.1f}  rule={r.get('rule')} ls={r.get('ls')} {match}  '{r['text'][:20]}'")
            results[fname] = dict(grid_pt=grid_pt, rows=rows)
    finally:
        word.Quit()

    OUT.write_text(json.dumps(results, indent=2, ensure_ascii=False, default=str))
    print(f"\nwrote {OUT}")


if __name__ == "__main__":
    main()
