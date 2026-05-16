"""COM-measure Word paragraph positions for L1-L8 line=0 atLeast repros.

For each repro: open in Word (visible=False), iterate paragraphs,
collect (start_y, end_y, sz). Per-paragraph advance gap = next.start_y -
this.start_y.

Output: tools/metrics/line0_atleast_word_measurements.json
"""
import os, json, glob
import win32com.client as win32

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
REPRO_DIR = os.path.join(REPO_ROOT, "tools", "metrics", "line0_atleast_repro")
OUT_PATH = os.path.join(REPO_ROOT, "tools", "metrics",
                        "line0_atleast_word_measurements.json")

wdVerticalPositionRelativeToPage = 6
wdActiveEndPageNumber = 3


def measure_one(word_app, path: str) -> dict:
    doc = word_app.Documents.Open(os.path.abspath(path), ReadOnly=True)
    try:
        paras = []
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            # R30 fix: collapsed start range
            r0 = doc.Range(rng.Start, rng.Start)
            y = r0.Information(wdVerticalPositionRelativeToPage)
            page = r0.Information(wdActiveEndPageNumber)
            text = (rng.Text or "").rstrip("\r\x07")
            # Get font size from first character if any
            sz = None
            if rng.Characters.Count > 0:
                try:
                    sz = float(rng.Characters(1).Font.Size)
                except Exception:
                    pass
            paras.append({
                "i": i, "y": y, "page": page, "text": text, "sz": sz,
            })
        # compute deltas
        for j in range(len(paras) - 1):
            if paras[j+1]["page"] == paras[j]["page"]:
                paras[j]["delta_to_next"] = round(paras[j+1]["y"] - paras[j]["y"], 3)
            else:
                paras[j]["delta_to_next"] = None
        return {"path": os.path.basename(path), "paragraphs": paras}
    finally:
        doc.Close(SaveChanges=False)


def main():
    word_app = win32.gencache.EnsureDispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = 0
    try:
        results = {}
        for fp in sorted(glob.glob(os.path.join(REPRO_DIR, "*.docx"))):
            name = os.path.splitext(os.path.basename(fp))[0]
            print(f"measuring {name}...")
            results[name] = measure_one(word_app, fp)
    finally:
        word_app.Quit()
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nwrote -> {OUT_PATH}")

    # Summary print: per-variant first few deltas
    for name, r in results.items():
        print(f"\n[{name}]")
        for p in r["paragraphs"][:5]:
            print(f"  wi={p['i']} sz={p['sz']:.1f}pt y={p['y']:.2f} page={p['page']} dy={p.get('delta_to_next')}")


if __name__ == "__main__":
    main()
