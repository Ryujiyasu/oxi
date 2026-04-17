"""Measure Word's per-paragraph Y inside Table 1 of d77a.

Hypothesis: Oxi collapses para spacing inside table cells, losing ~24pt/table.
If this is right, Word's 4 paras will have non-line-pitch gaps between them
(≈15.5pt within a para, but >15.5pt at paragraph boundaries).

Output: pipeline_data/d77a_t1_paraspacing.json + stderr summary.
"""
import os, sys, json, time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)
OUT = r"pipeline_data/d77a_t1_paraspacing.json"


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True)
        time.sleep(0.3)
        doc.Repaginate()

        tbl = doc.Tables(1)
        cell = tbl.Cell(1, 1)
        out = {"paragraphs": []}

        for pi, p in enumerate(cell.Range.Paragraphs, 1):
            pr = p.Range
            n = pr.Characters.Count
            # First-char Y (line 1 top)
            try:
                first_y = pr.Characters(1).Information(6)
            except Exception:
                first_y = None
            # Last-char Y (last line top)
            try:
                last_y = pr.Characters(n).Information(6) if n > 0 else first_y
            except Exception:
                last_y = None
            # Count distinct Y positions = lines
            ys = set()
            if n > 0:
                step = max(1, n // 60)
                for i in range(1, n + 1, step):
                    try:
                        ys.add(round(pr.Characters(i).Information(6), 1))
                    except Exception:
                        pass
                try:
                    ys.add(round(pr.Characters(n).Information(6), 1))
                except Exception:
                    pass
            else:
                if first_y:
                    ys.add(round(first_y, 1))

            sorted_ys = sorted(ys)
            text_head = pr.Text.replace("\r", "").replace("\x07", "")[:30]
            out["paragraphs"].append(
                {
                    "pi": pi,
                    "first_y": round(first_y, 2) if first_y else None,
                    "last_y": round(last_y, 2) if last_y else None,
                    "line_ys": sorted_ys,
                    "line_count": len(sorted_ys),
                    "text_head": text_head,
                }
            )

        # After-table Y (first content after table)
        tbl_end = tbl.Range.End
        try:
            after_y = doc.Range(tbl_end, tbl_end).Information(6)
        except Exception:
            after_y = None

        out["after_table_y"] = round(after_y, 2) if after_y else None

        # Decomposition
        print("=== Word d77a Table 1 per-paragraph ===", file=sys.stderr)
        for p in out["paragraphs"]:
            print(
                f"  para{p['pi']} lines={p['line_count']} "
                f"first_y={p['first_y']} last_y={p['last_y']} "
                f"text={p['text_head']!r}",
                file=sys.stderr,
            )
            if p["line_ys"]:
                diffs = [
                    round(p["line_ys"][i + 1] - p["line_ys"][i], 2)
                    for i in range(len(p["line_ys"]) - 1)
                ]
                print(f"    within-para gaps: {diffs}", file=sys.stderr)

        # Inter-para gaps
        print("\n=== Inter-paragraph gaps ===", file=sys.stderr)
        paras = out["paragraphs"]
        for i in range(len(paras) - 1):
            a = paras[i]
            b = paras[i + 1]
            if a["last_y"] and b["first_y"]:
                gap = round(b["first_y"] - a["last_y"], 2)
                print(
                    f"  p{a['pi']}->p{b['pi']} gap={gap}pt (last_y={a['last_y']} -> first_y={b['first_y']})",
                    file=sys.stderr,
                )

        # Total table content span
        if paras and paras[0]["first_y"] and paras[-1]["last_y"]:
            span = round(paras[-1]["last_y"] - paras[0]["first_y"], 2)
            print(f"\nContent span (para1.first_y -> paraN.last_y) = {span}pt", file=sys.stderr)
        if paras and paras[0]["first_y"] and after_y:
            total = round(after_y - paras[0]["first_y"], 2)
            print(f"Total to after-table = {total}pt", file=sys.stderr)

        os.makedirs("pipeline_data", exist_ok=True)
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print(f"\n[OK] {OUT}")

        doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
