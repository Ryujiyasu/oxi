"""Find the first Y-position divergence point between Word COM and Oxi
rendering for e3c545. Walks paragraphs in document order, computes
absolute Y position (taking page into account), and reports the first
paragraph where Oxi diverges from Word by > threshold.

Used as R7.34 instrumentation tool: shifts investigation from
outlier-chasing to systematic root-cause localization.

Output: list of (word_i, text_prefix, word_abs_y, oxi_abs_y, delta) sorted
by word_i, plus identification of the first significant divergence.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent.parent

DOC_ID = "e3c545"
DOCX = ROOT / "tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx"
WORD_DATA = ROOT / "pipeline_data/cascade_word_y" / f"{DOC_ID}.json"
RENDERER = ROOT / "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"

# Page height (A4 = 841.9pt), used to compute absolute Y
PAGE_H = 841.9
DIVERGENCE_THRESHOLD = 5.0  # pt — flag anything beyond this


def abs_y(page: int, y: float) -> float:
    return (page - 1) * PAGE_H + y


def main():
    # Load Word COM data
    with open(WORD_DATA, encoding="utf-8") as f:
        word = json.load(f)
    word_paras = word["paragraphs"]
    print(f"Word: {len(word_paras)} paragraphs across {word['n_pages']} pages")

    # Render Oxi and get layout dump
    dump_path = Path(r"C:/Users/ryuji/AppData/Local/Temp") / f"{DOC_ID}_ydiv.json"
    cmd = [str(RENDERER), str(DOCX), str(dump_path.with_suffix("")) + "_",
           f"--dump-layout={dump_path}"]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        print(f"Renderer failed: {result.stderr[:300]}", file=sys.stderr)
        return 2
    with open(dump_path, encoding="utf-8") as f:
        oxi = json.load(f)
    print(f"Oxi: {len(oxi['pages'])} pages")

    # Build Oxi: list of (page, abs_y, text_normalized_prefix) per text element
    # Grouped by (page, para_idx, cell_para_idx, y) → take first text fragment
    oxi_lines = []  # (abs_y, text_prefix, page, para_idx, cpi)
    for pg in oxi["pages"]:
        page_num = pg["page"]
        # Group by (para_idx, cpi, rounded y)
        lines: dict = {}
        for el in pg.get("elements", []):
            if el.get("type") != "text":
                continue
            pi = el.get("para_idx")
            cpi = el.get("cell_para_idx")
            y_key = round(el["y"], 1)
            key = (pi, cpi, y_key)
            slot = lines.setdefault(key, {"y": el["y"], "x_min": el["x"], "parts": []})
            slot["parts"].append((el["x"], el["text"]))
            if el["x"] < slot["x_min"]:
                slot["x_min"] = el["x"]
        # Build line records
        for (pi, cpi, _), slot in lines.items():
            slot["parts"].sort(key=lambda xt: xt[0])
            text = "".join(t for _, t in slot["parts"])
            oxi_lines.append({
                "page": page_num,
                "abs_y": abs_y(page_num, slot["y"]),
                "y_local": slot["y"],
                "text": text,
                "pi": pi,
                "cpi": cpi,
            })

    # Sort Oxi lines by abs_y (document order)
    oxi_lines.sort(key=lambda r: r["abs_y"])
    print(f"Oxi: {len(oxi_lines)} distinct text lines")

    # Match each Word paragraph to an Oxi line by text prefix
    # Word paragraphs may have empty text; skip those
    matches = []
    oxi_used = [False] * len(oxi_lines)
    for wp in word_paras:
        wt = (wp.get("text") or "").strip()
        if len(wt) < 3:
            continue
        wt_prefix = wt[:25]
        wpage = wp.get("page")
        wy = wp.get("y_pt") or 0
        w_abs = abs_y(wpage, wy)
        # Search for matching Oxi line with same prefix, prefer closest abs_y
        best_idx = None
        best_dist = 9e9
        for i, ol in enumerate(oxi_lines):
            if oxi_used[i]:
                continue
            ot = ol["text"][:25]
            if ot == wt_prefix or (len(wt_prefix) >= 6 and ot.startswith(wt_prefix[:6])):
                dist = abs(ol["abs_y"] - w_abs)
                if dist < best_dist:
                    best_dist = dist
                    best_idx = i
        if best_idx is not None:
            oxi_used[best_idx] = True
            ol = oxi_lines[best_idx]
            delta = ol["abs_y"] - w_abs
            matches.append({
                "word_i": wp["i"],
                "word_page": wpage,
                "word_y": wy,
                "word_abs_y": w_abs,
                "oxi_page": ol["page"],
                "oxi_y": ol["y_local"],
                "oxi_abs_y": ol["abs_y"],
                "delta": delta,
                "text": wt_prefix.encode("cp932", errors="replace").decode("cp932"),
            })

    matches.sort(key=lambda m: m["word_i"])
    print(f"Matched: {len(matches)} of {len(word_paras)} Word paragraphs")

    # Find first divergence
    print("\nFirst 10 matches:")
    for m in matches[:10]:
        delta_str = f"{m['delta']:+.1f}"
        print(f"  w_i={m['word_i']:4d} wp={m['word_page']} op={m['oxi_page']} "
              f"wy={m['word_y']:.1f} oy={m['oxi_y']:.1f} delta={delta_str}pt")

    # First divergence > threshold
    first_div = None
    prev = None
    for m in matches:
        if prev is None:
            prev = m["delta"]
            first_div = None
            continue
        # Detect jump from prev delta — magnitude > threshold
        if abs(m["delta"] - prev) > DIVERGENCE_THRESHOLD and first_div is None:
            first_div = m
            first_div["prev_delta"] = prev
            break
        prev = m["delta"]

    if first_div:
        print(f"\nFIRST significant Y divergence:")
        print(f"  word_i={first_div['word_i']}  text={first_div['text']!r}")
        print(f"  word page={first_div['word_page']} y={first_div['word_y']:.1f}")
        print(f"  oxi  page={first_div['oxi_page']} y={first_div['oxi_y']:.1f}")
        print(f"  delta jumped from {first_div['prev_delta']:+.1f} to {first_div['delta']:+.1f} pt")
    else:
        print(f"\nNo divergence > {DIVERGENCE_THRESHOLD}pt found")

    # Show running delta trajectory
    print("\nDelta trajectory (every 30 matched paragraphs):")
    for i in range(0, len(matches), 30):
        m = matches[i]
        print(f"  w_i={m['word_i']:4d}  delta={m['delta']:+7.1f}pt  text={m['text'][:30]!r}")

    # Output JSON
    out_path = ROOT / "pipeline_data" / f"{DOC_ID}_y_divergence.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({
            "doc_id": DOC_ID,
            "n_word_paras": len(word_paras),
            "n_matched": len(matches),
            "first_divergence": first_div,
            "matches": matches,
        }, f, ensure_ascii=False, indent=2)
    print(f"\nFull data → {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
