"""Measure border overhead with different border configurations.

Tests:
1. No borders at all
2. Outer borders only (no insideH/insideV)
3. InsideH only (no outer)
4. All borders
5. Top border only on specific rows
"""
import win32com.client
import json

def measure():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        results = {}

        configs = {
            "no_border": {"outer": False, "insideH": False, "insideV": False},
            "outer_only": {"outer": True, "insideH": False, "insideV": False},
            "insideH_only": {"outer": False, "insideH": True, "insideV": False},
            "all_borders": {"outer": True, "insideH": True, "insideV": True},
        }

        for config_name, cfg in configs.items():
            for bw_eighths in [4, 8, 12]:  # 0.5pt, 1pt, 1.5pt
                bw_pt = bw_eighths / 8.0
                doc = word.Documents.Add()

                sel = word.Selection
                rng = sel.Range
                tbl = doc.Tables.Add(rng, 5, 2)
                tbl.Range.Font.Name = "Calibri"
                tbl.Range.Font.Size = 11

                # Clear all borders first
                for bid in [-1, -2, -3, -4, -5, -6]:
                    try:
                        tbl.Borders(bid).LineStyle = 0
                    except:
                        pass

                # Set requested borders
                if cfg["outer"]:
                    for bid in [-1, -2, -3, -4]:  # top, left, bottom, right
                        tbl.Borders(bid).LineStyle = 1
                        tbl.Borders(bid).LineWidth = bw_eighths
                if cfg["insideH"]:
                    tbl.Borders(-5).LineStyle = 1  # insideH
                    tbl.Borders(-5).LineWidth = bw_eighths
                if cfg["insideV"]:
                    tbl.Borders(-6).LineStyle = 1
                    tbl.Borders(-6).LineWidth = bw_eighths

                # Add text
                for ri in range(1, 6):
                    for ci in range(1, 3):
                        tbl.Cell(ri, ci).Range.Text = "A"

                # Measure
                ys = []
                for ri in range(1, 6):
                    y = tbl.Cell(ri, 1).Range.Paragraphs(1).Range.Information(6)
                    ys.append(round(y, 2))

                heights = [round(ys[i+1] - ys[i], 2) for i in range(len(ys)-1)]

                key = f"{config_name}_bw{bw_pt}"
                results[key] = {
                    "config": config_name,
                    "bw": bw_pt,
                    "ys": ys,
                    "heights": heights,
                }
                # Heights: row1=first, row2-3=middle, row4=last
                print(f"{key:30s}  h={heights}  (first={heights[0]}, mid={heights[1]}, last={heights[3]})")

                doc.Close(False)

        return results
    finally:
        word.Quit()


if __name__ == "__main__":
    results = measure()
    print("\n=== Border Overhead Analysis ===")
    print(f"{'Config':30s}  {'first':>6s}  {'mid2':>6s}  {'mid3':>6s}  {'last':>6s}  overhead_per_row")
    base_h = 18.0  # Calibri 11pt no-border
    for key in sorted(results.keys()):
        r = results[key]
        hs = r["heights"]
        overhead = [round(h - base_h, 2) for h in hs]
        print(f"{key:30s}  {hs[0]:6.1f}  {hs[1]:6.1f}  {hs[2]:6.1f}  {hs[3]:6.1f}  {overhead}")
