"""Measure Word's rendering of OMML fixtures.

For each fixture docx, opens in Word COM and measures:
- Math paragraph top y (via Paragraphs(2).Range.Information(6))
- Next paragraph top y (height consumed by the math)
- Font names detected within the math (expect Cambria Math)
- Optional per-char measurement of key elements

Output: tools/metrics/output/omml_fixtures_measurements.json
"""
import json, time, sys
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FIXTURES_DIR = Path(__file__).resolve().parent.parent / "fixtures" / "omml_samples"
OUT = Path(__file__).with_name("output") / "omml_fixtures_measurements.json"
OUT.parent.mkdir(parents=True, exist_ok=True)


def measure(word, docx: Path):
    try:
        doc = word.Documents.Open(str(docx.resolve()), ReadOnly=True)
        time.sleep(0.5)
        n_paras = doc.Paragraphs.Count
        # Expected: 3 paragraphs (label, math, END)
        para_ys = []
        for i in range(1, min(n_paras + 1, 5)):
            p = doc.Paragraphs(i)
            try:
                y = p.Range.Information(6)
                x = p.Range.Information(5)
                text = p.Range.Text[:40]
                font = p.Range.Font.Name
                size = p.Range.Font.Size
                para_ys.append({
                    "idx": i, "y": round(y, 2), "x": round(x, 2),
                    "text": text, "font": font, "size": size,
                })
            except Exception as e:
                para_ys.append({"idx": i, "error": str(e)})

        # Math paragraph height = y[2] - y[1]? Or y[3] - y[2]?
        # Structure: label (p1), math (p2), END (p3). So math_h = y[3] - y[2].
        math_h = None
        if len(para_ys) >= 3 and "y" in para_ys[1] and "y" in para_ys[2]:
            math_h = round(para_ys[2]["y"] - para_ys[1]["y"], 2)

        # Try to measure individual OMath objects
        omath_info = []
        try:
            if hasattr(doc, "OMaths"):
                for i in range(1, doc.OMaths.Count + 1):
                    om = doc.OMaths(i)
                    omath_info.append({
                        "idx": i,
                        "type": om.Type,  # 0=inline, 1=display
                        "range_text": om.Range.Text[:40],
                    })
        except Exception as e:
            omath_info = [{"error": str(e)}]

        doc.Close(False)
        return {
            "file": docx.name,
            "n_paragraphs": n_paras,
            "para_ys": para_ys,
            "math_h": math_h,
            "omath_info": omath_info,
        }
    except Exception as e:
        return {"file": docx.name, "error": str(e)}


def main():
    if not FIXTURES_DIR.exists():
        print(f"Fixtures dir not found: {FIXTURES_DIR}")
        return

    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False

    results = []
    try:
        for docx in sorted(FIXTURES_DIR.glob("*.docx")):
            print(f"\n=== {docx.name} ===")
            r = measure(word, docx)
            if "error" in r:
                print(f"  ERR: {r['error']}")
            else:
                for p in r["para_ys"]:
                    if "y" in p:
                        print(f"  p{p['idx']}: y={p['y']:>7.2f} font={p['font']} size={p['size']} text={p['text']!r}")
                if r["math_h"] is not None:
                    print(f"  math height: {r['math_h']:.2f}pt")
                for om in r["omath_info"]:
                    if "error" in om:
                        print(f"  OMath enum err: {om['error']}")
                    else:
                        print(f"  OMath {om['idx']}: type={om['type']} text={om['range_text']!r}")
            results.append(r)
    finally:
        try: word.Quit()
        except: pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"\n=== Summary ===")
    heights = [r.get("math_h") for r in results if r.get("math_h")]
    print(f"Measured {len(heights)}/{len(results)} fixtures")
    if heights:
        print(f"Math heights: min={min(heights):.2f}, max={max(heights):.2f}")
    print(f"Saved → {OUT}")


if __name__ == "__main__":
    main()
