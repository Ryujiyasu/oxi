"""Measure b35 row 6/7 content para positions to localize the +8/-9pt drift.

Row 6 Word_h=20.45pt (1 content line fs=10.5). Row 7 Word_h=55.55pt (multi-line).
Oxi: row 6=28.62 row 7=46.12. Border drawn 8pt lower in Oxi.

This script measures Word paragraph Y for each para in rows 1-8 to see where
the Word border actually lands vs Oxi.
"""
import json, time, sys
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path(r"C:/Users/ryuji/oxi-4/tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx")

word = win32com.client.Dispatch("Word.Application")
time.sleep(1.0)
word.Visible = False
word.DisplayAlerts = False

try:
    doc = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
    time.sleep(1.5)

    # Iterate all paragraphs on page 1
    n = doc.Paragraphs.Count
    print(f"Total paragraphs: {n}")
    p1_paras = []
    for i in range(1, n + 1):
        p = doc.Paragraphs(i)
        rng = p.Range
        try:
            pg = int(rng.Information(3))
            if pg > 1: break  # stop after p1
            y = rng.Information(6)
            x = rng.Information(5)
        except Exception:
            continue
        text = rng.Text[:40].replace('\r', '\\r')
        # Detect paragraph font/size
        fs = rng.Font.Size
        font = rng.Font.Name
        # Is paragraph inside a table?
        try:
            in_tbl = rng.Information(12)  # wdWithInTable
        except:
            in_tbl = None
        p1_paras.append({"idx": i, "page": pg, "y": round(y, 2), "x": round(x, 2), "fs": fs, "font": font, "in_table": bool(in_tbl), "text": text})

    doc.Close(False)

    print(f"\n{len(p1_paras)} paragraphs on p1:")
    for p in p1_paras:
        flag = 'T' if p["in_table"] else ' '
        print(f"  idx={p['idx']:>3} {flag} y={p['y']:>7.2f} fs={p['fs']:>5} font={p['font'][:12]:<12} text={p['text']!r}")

    # Save
    OUT = Path(__file__).with_name("output") / "b35_p1_para_ys.json"
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(p1_paras, f, ensure_ascii=False, indent=2)
    print(f"\nSaved → {OUT}")

finally:
    try: word.Quit()
    except: pass
