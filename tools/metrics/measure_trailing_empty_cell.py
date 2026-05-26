"""Measure Word vs Oxi for the trailing-empty-cell repro fixtures (S304).

Hypothesis check: Δy of "AFTER TABLE marker." between v1 (no trailing
empty) and v2 (1 trailing empty <w:p sz=18>) should be ~12.5pt in Word
and ~0 in Oxi if the bug exists.
"""
import json
import os
import subprocess
import sys

import win32com.client as w32

REPRO_DIR = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "fixtures", "trailing_empty_cell_repro"))
GDI = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "oxi-gdi-renderer", "target",
    "release", "oxi-gdi-renderer.exe"))


def measure_word(docx_path: str) -> dict:
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        out = {"path": docx_path, "paras": []}
        for pi in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            rng_start = doc.Range(rng.Start, rng.Start)
            text = rng.Text[:60].replace("\r", "\\r").replace("\x07", "\\x07")
            try:
                y = rng_start.Information(6)
                pg = rng_start.Information(3)
            except Exception:
                y, pg = None, None
            in_tbl = rng.Information(12)
            out["paras"].append({"i": pi, "page": pg, "y": y,
                                  "in_tbl": bool(in_tbl), "text": text})
        doc.Close(False)
        return out
    finally:
        word.Quit()


def measure_oxi(docx_path: str) -> dict:
    out_prefix = os.path.join(os.environ.get("TEMP", "/tmp"), "_te_out")
    layout = os.path.join(os.environ.get("TEMP", "/tmp"),
                          os.path.basename(docx_path) + ".layout.json")
    subprocess.run([GDI, docx_path, out_prefix, "150",
                    f"--dump-layout={layout}"], check=True, capture_output=True)
    with open(layout, encoding="utf-8") as f:
        d = json.load(f)
    # gather all text elements with text content, group by y
    paras = []
    for pi, page in enumerate(d["pages"], 1):
        text_elems = [e for e in page["elements"] if e.get("type") == "text"]
        text_elems.sort(key=lambda e: (e["y"], e["x"]))
        # Group adjacent same-y elements as a line
        from collections import defaultdict
        by_y = defaultdict(list)
        for e in text_elems:
            yk = round(e["y"], 1)
            by_y[yk].append(e)
        for yk in sorted(by_y):
            es = sorted(by_y[yk], key=lambda e: e["x"])
            txt = "".join((e.get("text", "") or "") for e in es)[:60]
            paras.append({"page": pi, "y": yk, "text": txt})
    return {"path": docx_path, "lines": paras}


def find_after_marker(word_paras, oxi_lines):
    w_after = None
    for p in word_paras:
        if "AFTER TABLE marker" in p["text"]:
            w_after = p
            break
    o_after = None
    for l in oxi_lines:
        if "AFTER TABLE" in l["text"]:
            o_after = l
            break
    return w_after, o_after


def main():
    fixtures = ["v1_text_only.docx", "v2_text_plus_empty.docx",
                "v3_text_plus_2empty.docx", "v4_only_empty.docx"]
    results = {}
    for f in fixtures:
        path = os.path.join(REPRO_DIR, f)
        if not os.path.exists(path):
            print(f"  SKIP {f}: not found")
            continue
        print(f"--- {f} ---")
        wr = measure_word(path)
        oxi = measure_oxi(path)
        w_after, o_after = find_after_marker(wr["paras"], oxi["lines"])
        if w_after and o_after:
            dy = o_after["y"] - w_after["y"]
            print(f"  Word AFTER: page={w_after['page']} y={w_after['y']:.2f}")
            print(f"  Oxi  AFTER: page={o_after['page']} y={o_after['y']:.2f}")
            print(f"  dy (oxi - word) = {dy:+.2f}pt")
            results[f] = {"word_y": w_after["y"], "oxi_y": o_after["y"],
                          "dy": dy}
        else:
            print(f"  MISSING marker w={w_after is not None} o={o_after is not None}")
            results[f] = None

    print()
    print("=== Summary (dy of 'AFTER TABLE marker.') ===")
    print(f"{'variant':<28} {'word_y':>8} {'oxi_y':>8} {'dy':>7}")
    for f in fixtures:
        r = results.get(f)
        if r:
            print(f"  {f:<26} {r['word_y']:>8.2f} {r['oxi_y']:>8.2f} {r['dy']:>+7.2f}")

    print()
    if all(results.get(v) for v in ["v1_text_only.docx", "v2_text_plus_empty.docx"]):
        word_delta_v1_v2 = (results["v2_text_plus_empty.docx"]["word_y"]
                            - results["v1_text_only.docx"]["word_y"])
        oxi_delta_v1_v2 = (results["v2_text_plus_empty.docx"]["oxi_y"]
                           - results["v1_text_only.docx"]["oxi_y"])
        print(f"Word reserves for 1 trailing empty: {word_delta_v1_v2:+.2f}pt")
        print(f"Oxi  reserves for 1 trailing empty: {oxi_delta_v1_v2:+.2f}pt")
        print(f"Gap (Word - Oxi)                  : {word_delta_v1_v2 - oxi_delta_v1_v2:+.2f}pt")

    with open(os.path.join(os.path.dirname(__file__),
                            "trailing_empty_cell_measurement.json"), "w") as f:
        json.dump(results, f, indent=2)


if __name__ == "__main__":
    sys.exit(main())
