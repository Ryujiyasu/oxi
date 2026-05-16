"""COM-measure paragraph Y for M1-M6 LM2 cursor repros."""
import os, json, glob
import win32com.client as win32

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
REPRO_DIR = os.path.join(REPO_ROOT, "tools", "metrics", "lm2_cursor_repro")
OUT_PATH = os.path.join(REPO_ROOT, "tools", "metrics",
                        "lm2_cursor_word_measurements.json")

wdVerticalPositionRelativeToPage = 6


def measure_one(word_app, path):
    doc = word_app.Documents.Open(os.path.abspath(path), ReadOnly=True)
    try:
        paras = []
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            r0 = doc.Range(rng.Start, rng.Start)
            y = r0.Information(wdVerticalPositionRelativeToPage)
            page = r0.Information(3)
            text = (rng.Text or "").rstrip("\r\x07")
            paras.append({"i": i, "y": y, "page": page, "text": text})
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
    print(f"\n-> {OUT_PATH}\n")
    for name, r in results.items():
        print(f"[{name}]")
        for p in r["paragraphs"]:
            kind = "TEXT" if p["text"].strip() else "EMPTY"
            print(f"  wi={p['i']} {kind} y={p['y']:.2f} dy={p.get('delta_to_next')}")
        print()


if __name__ == "__main__":
    main()
