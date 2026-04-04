"""COM measurement: paragraph spacing under grid snap.

Target: gen2_023 (docDefaults sa=200tw, line=276, docGrid linePitch=360)
Goal: Determine exact spacing values when grid snap is active.

Measures:
1. All paragraph Y coordinates (wdVerticalPositionRelativeToPage)
2. Inter-paragraph gaps (Y[n+1] - Y[n])
3. Style info per paragraph
"""
import win32com.client
import os
import sys
import json
import time

def measure_doc(docx_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone

    abs_path = os.path.abspath(docx_path)
    print(f"Opening: {abs_path}")
    doc = word.Documents.Open(abs_path, ReadOnly=True)
    time.sleep(1)

    results = {
        "file": os.path.basename(docx_path),
        "paragraphs": [],
    }

    n_paras = doc.Paragraphs.Count
    print(f"Total paragraphs: {n_paras}")

    for i in range(1, n_paras + 1):
        para = doc.Paragraphs(i)
        rng = para.Range

        # Y coordinate (wdVerticalPositionRelativeToPage = 6)
        y = rng.Information(6)
        # Page number (wdActiveEndPageNumber = 3)
        page = rng.Information(3)

        # Style info
        style_name = para.Style.NameLocal
        fmt = para.Format
        sa = fmt.SpaceAfter          # points
        sb = fmt.SpaceBefore         # points
        ls = fmt.LineSpacing         # points (setting value)
        lr = fmt.LineSpacingRule     # 0=auto, 1=atLeast, 2=exactly, 4=multiple, etc
        snap = not fmt.NoLineNumber  # proxy — not reliable

        text = rng.Text[:40].replace('\r', '\\r').replace('\n', '\\n')

        entry = {
            "index": i,
            "page": page,
            "y": round(y, 2),
            "style": style_name,
            "sa": round(sa, 2),
            "sb": round(sb, 2),
            "ls": round(ls, 2),
            "lr": lr,
            "text": text,
        }
        results["paragraphs"].append(entry)

        if i <= 60 or page >= 2:
            print(f"  P{i:3d} page={page} y={y:7.1f} sa={sa:5.1f} sb={sb:5.1f} ls={ls:5.1f} lr={lr} [{style_name}] {text[:30]}")

    # Compute gaps
    print("\n=== GAPS ===")
    paras = results["paragraphs"]
    for i in range(1, len(paras)):
        prev = paras[i-1]
        cur = paras[i]
        if prev["page"] == cur["page"]:
            gap = round(cur["y"] - prev["y"], 2)
            print(f"  P{prev['index']:3d}→P{cur['index']:3d}: gap={gap:6.2f}pt  (sa={prev['sa']}, sb={cur['sb']}, ls={prev['ls']})")

    doc.Close(False)
    word.Quit()

    return results


if __name__ == "__main__":
    if len(sys.argv) < 2:
        # Default to gen2_023
        docx = os.path.join(os.path.dirname(__file__), "..", "..",
                           "tools", "golden-test", "documents", "docx",
                           "gen2_023_育児休業規程.docx")
    else:
        docx = sys.argv[1]

    results = measure_doc(docx)

    # Save results
    out_path = os.path.join(os.path.dirname(__file__), "..", "..",
                            "pipeline_data", "ra_manual_measurements.json")
    # Append to existing
    existing = []
    if os.path.exists(out_path):
        with open(out_path, encoding="utf-8") as f:
            try:
                existing = json.load(f)
                if not isinstance(existing, list):
                    existing = [existing]
            except:
                existing = []

    existing.append({
        "measurement": "spacing_grid_snap",
        "date": "2026-04-03",
        "results": results
    })

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {out_path}")
