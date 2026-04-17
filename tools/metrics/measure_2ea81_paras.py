"""Measure Word paragraph positions for 2ea81 p2 baseline."""
import json, time, sys
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path(r"C:/Users/ryuji/oxi-4/tools/golden-test/documents/docx/2ea81a8441cc_0025006-192.docx")
OUT = Path(__file__).with_name("output") / "2ea81_paras.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

word = win32com.client.Dispatch("Word.Application")
time.sleep(1.0)
word.Visible = False
word.DisplayAlerts = False

try:
    doc = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
    time.sleep(1.5)

    paras = []
    n = doc.Paragraphs.Count
    print(f"Total paragraphs: {n}")
    for i in range(1, n + 1):
        p = doc.Paragraphs(i)
        rng = p.Range
        try:
            y = rng.Information(6)
            x = rng.Information(5)
            pg = int(rng.Information(3))
        except Exception:
            continue
        text = rng.Text[:40]
        paras.append({"idx": i, "page": pg, "y": round(y, 2), "x": round(x, 2), "chars": len(rng.Text), "text": text})

    doc.Close(False)

    # Group by page
    from collections import defaultdict
    by_page = defaultdict(list)
    for p in paras:
        by_page[p["page"]].append(p)

    print(f"\nPages found: {sorted(by_page.keys())}")
    for pg in sorted(by_page.keys()):
        plist = by_page[pg]
        print(f"\n=== Page {pg} ({len(plist)} paragraphs, y={plist[0]['y']:.1f}..{plist[-1]['y']:.1f}) ===")
        for p in plist[:3]:
            print(f"  idx={p['idx']:>3} y={p['y']:>7.2f} text={p['text']!r}")
        if len(plist) > 6:
            print(f"  ... ({len(plist)-6} more) ...")
        for p in plist[-3:]:
            print(f"  idx={p['idx']:>3} y={p['y']:>7.2f} text={p['text']!r}")

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(paras, f, ensure_ascii=False, indent=2)
    print(f"\nSaved → {OUT}")

finally:
    try: word.Quit()
    except: pass
