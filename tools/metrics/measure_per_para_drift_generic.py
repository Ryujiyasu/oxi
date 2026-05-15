"""Per-paragraph y-drift analysis for any baseline doc (generic version).

Generalizes `measure_3a4f9f_per_para_drift.py` to any doc_id by extracting
page geometry from the doc's OOXML sectPr. Used to confirm whether
ed025c's Phase 1 outlier shares the same drift-jump mechanism as 3a4f9f
(documented in [[session59-3a4f9f-drift-jumps-floating-table-footprint]]).

Usage:
  python tools/metrics/measure_per_para_drift_generic.py <doc_id_prefix>
  e.g.:
  python tools/metrics/measure_per_para_drift_generic.py ed025c

Output: pipeline_data/ra_manual_measurements/<doc_id>_per_para_drift.json
Also prints a "jump" trace highlighting Δdrift > 100pt between consecutive
matched paragraphs (the diagnostic that surfaced 3a4f9f's 3 jump events).

Instrumentation only — does NOT modify oxidocs-core or change any baseline.
Phase 1 53/55 mean 0.9842 must remain unchanged.
"""
from __future__ import annotations

import json
import os
import re
import statistics
import sys
import zipfile
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = r"c:\Users\ryuji\oxi-main"
DOCS_DIR = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
WORD_SUMMARY = os.path.join(REPO, "pipeline_data", "pagination_word", "_summary.json")

# Twip → point conversion: 1pt = 20 twips
TW2PT = 1.0 / 20.0

# Paragraph type classification — based on Word text. Order matters.
TYPE_PATTERNS = [
    ("empty",           re.compile(r"^[\s　]*$")),
    ("page_number",     re.compile(r"^[\s　]*\d{1,3}[\s　]*$")),
    ("chapter_kanji",   re.compile(r"^[\s　]*第[一二三四五六七八九十百\d０-９]+章")),
    ("article_kanji",   re.compile(r"^[\s　]*第[一二三四五六七八九十百\d０-９]+条")),
    ("bracket_pair",    re.compile(r"^[\s　]*[【〔]")),
    ("list_marker",     re.compile(r"^[\s　]*[・]")),
    ("numbered_paren",  re.compile(r"^[\s　]*[（\(][一二三四五六七八九十\d０-９]")),
    ("numbered_kanji",  re.compile(r"^[\s　]*[一二三四五六七八九十][\s　]")),
    ("numbered_arabic", re.compile(r"^[\s　]*[\d０-９]+[\s　]")),
    ("numbered_circle", re.compile(r"^[\s　]*[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮]")),
]


def classify(text: str, in_table: bool) -> str:
    t = text or ""
    base = "body"
    for label, pat in TYPE_PATTERNS:
        if pat.search(t):
            base = label
            break
    if in_table:
        return f"{base}_intable"
    return base


def resolve_doc(prefix: str) -> tuple[str, str]:
    """Resolve a doc_id prefix to (full_doc_id, docx_filename) via pagination_word summary."""
    with open(WORD_SUMMARY, encoding="utf-8") as f:
        data = json.load(f)
    matches = [d for d in data["docs"] if d["doc_id"].startswith(prefix)]
    if not matches:
        raise SystemExit(f"No doc_id matching prefix '{prefix}'")
    if len(matches) > 1:
        raise SystemExit(f"Ambiguous prefix '{prefix}': {[d['doc_id'] for d in matches]}")
    return matches[0]["doc_id"], matches[0]["filename"]


def extract_geometry(docx_path: str) -> dict:
    """Pull pgSz/pgMar from the FIRST sectPr in word/document.xml.

    Word may have multiple sections; the first sectPr applies to the document
    body. For our pagination-drift purposes, the body geometry is what matters.
    """
    with zipfile.ZipFile(docx_path) as z, z.open("word/document.xml") as fh:
        xml = fh.read().decode("utf-8", errors="replace")
    sect_m = re.search(r"<w:sectPr\b.*?</w:sectPr>", xml, re.DOTALL)
    if not sect_m:
        raise SystemExit("No <w:sectPr> found in word/document.xml")
    s = sect_m.group(0)
    pgsz = re.search(r'<w:pgSz\s+([^/]*?)/?>', s)
    pgmar = re.search(r'<w:pgMar\s+([^/]*?)/?>', s)
    if not pgsz or not pgmar:
        raise SystemExit("Missing pgSz/pgMar in sectPr")

    def attrs(s: str) -> dict:
        return dict(re.findall(r'w:(\w+)="([^"]*)"', s))

    sz = attrs(pgsz.group(1))
    mar = attrs(pgmar.group(1))
    return {
        "page_w_tw": int(sz["w"]),
        "page_h_tw": int(sz["h"]),
        "top_tw": int(mar["top"]),
        "bottom_tw": int(mar["bottom"]),
        "left_tw": int(mar["left"]),
        "right_tw": int(mar["right"]),
        "page_h_pt": int(sz["h"]) * TW2PT,
        "top_pt": int(mar["top"]) * TW2PT,
        "bottom_pt": int(mar["bottom"]) * TW2PT,
        "content_h_pt": (int(sz["h"]) - int(mar["top"]) - int(mar["bottom"])) * TW2PT,
    }


def linear_y(page: int | None, y: float | None, top_margin: float, content_h: float) -> float | None:
    if page is None or y is None or page < 1:
        return None
    return (page - 1) * content_h + (y - top_margin)


def main():
    if len(sys.argv) < 2:
        raise SystemExit("Usage: measure_per_para_drift_generic.py <doc_id_prefix>")
    prefix = sys.argv[1]
    doc_id, fname = resolve_doc(prefix)
    docx_path = os.path.join(DOCS_DIR, fname)
    geom = extract_geometry(docx_path)
    top = geom["top_pt"]
    content_h = geom["content_h_pt"]

    diff_path = os.path.join(REPO, "pipeline_data", "pagination_diff", f"{doc_id}.json")
    word_path = os.path.join(REPO, "pipeline_data", "pagination_word", f"{doc_id}.json")
    oxi_path = os.path.join(REPO, "pipeline_data", "pagination_oxi", f"{doc_id}.json")
    out_path = os.path.join(
        REPO, "pipeline_data", "ra_manual_measurements",
        f"{doc_id}_per_para_drift.json",
    )

    with open(diff_path, encoding="utf-8") as f:
        diff = json.load(f)
    with open(word_path, encoding="utf-8") as f:
        word = json.load(f)
    with open(oxi_path, encoding="utf-8") as f:
        oxi = json.load(f)

    print(f"=== Per-paragraph drift: {doc_id} ===")
    print(f"  Filename: {fname}")
    print(f"  Geometry: page_h={geom['page_h_pt']:.2f}pt, top={top:.2f}pt, "
          f"bottom={geom['bottom_pt']:.2f}pt, content_h={content_h:.2f}pt")
    print(f"  n_matches: {len(diff['matches'])}")

    word_by_i = {p["i"]: p for p in word["paragraphs"]}
    oxi_by_page_text = defaultdict(list)
    for page_num, entries in oxi["pages"].items():
        pg = int(page_num)
        for e in entries:
            t = e.get("text") or ""
            oxi_by_page_text[(pg, t)].append({
                "y": e.get("y"),
                "para_idx": e.get("para_idx"),
            })

    paragraphs = []
    skipped = defaultdict(int)
    for seq_idx, m in enumerate(diff["matches"]):
        wi = m.get("word_i")
        wp = m.get("word_page")
        op = m.get("oxi_page")
        text = m.get("text") or ""
        delta_p = m.get("page_delta")

        wpara = word_by_i.get(wi)
        if not wpara:
            skipped["no_word_para"] += 1
            continue
        in_table = bool(wpara.get("in_table", False))
        word_y = wpara.get("y")
        word_lin = linear_y(wp, word_y, top, content_h)

        oxi_lin = None
        oxi_y = None
        if op is not None:
            cands = oxi_by_page_text.get((op, text), [])
            if len(cands) == 1:
                oxi_y = cands[0]["y"]
                oxi_lin = linear_y(op, oxi_y, top, content_h)
            elif len(cands) > 1:
                if word_lin is not None:
                    expected_oxi_lin = word_lin
                    best = min(cands, key=lambda c: abs(
                        (linear_y(op, c["y"], top, content_h) or 0) - expected_oxi_lin))
                    oxi_y = best["y"]
                    oxi_lin = linear_y(op, oxi_y, top, content_h)
                    skipped["multi_oxi_resolved_by_distance"] += 1
                else:
                    oxi_y = cands[0]["y"]
                    oxi_lin = linear_y(op, oxi_y, top, content_h)
                    skipped["multi_oxi_first"] += 1
            else:
                skipped["no_oxi_text_match"] += 1

        ptype = classify(text, in_table)
        drift = None
        if word_lin is not None and oxi_lin is not None:
            drift = round(oxi_lin - word_lin, 3)

        paragraphs.append({
            "seq_idx": seq_idx,
            "word_i": wi,
            "word_page": wp,
            "word_y": word_y,
            "oxi_page": op,
            "oxi_y": oxi_y,
            "word_linear": round(word_lin, 2) if word_lin is not None else None,
            "oxi_linear": round(oxi_lin, 2) if oxi_lin is not None else None,
            "drift": drift,
            "page_delta": delta_p,
            "type": ptype,
            "text_preview": text[:30],
            "in_table": in_table,
        })

    valid = [p for p in paragraphs if p["drift"] is not None]
    print(f"  Records with valid drift: {len(valid)}")
    print(f"  Skipped: {dict(skipped)}")
    if not valid:
        return

    valid_sorted = sorted(valid, key=lambda p: p["word_i"])

    # JUMP detection: |Δdrift| > 100pt between consecutive matched paragraphs
    print()
    print("=== JUMP events (|Δdrift| > 100pt between consecutive matched paragraphs) ===")
    print(f'{"seq":>5}  {"word_i_prev->cur":>18}  {"wp":>3}  {"op":>3}  {"Δp":>4}  '
          f'{"drift_prev":>11}  {"drift_cur":>11}  {"Δdrift":>9}  text_preview')
    prev = None
    jumps = []
    for p in valid_sorted:
        if prev is not None:
            dd = round(p["drift"] - prev["drift"], 2)
            if abs(dd) > 100:
                jumps.append({
                    "prev_word_i": prev["word_i"],
                    "cur_word_i": p["word_i"],
                    "prev_word_page": prev["word_page"],
                    "cur_word_page": p["word_page"],
                    "prev_oxi_page": prev["oxi_page"],
                    "cur_oxi_page": p["oxi_page"],
                    "prev_drift": prev["drift"],
                    "cur_drift": p["drift"],
                    "delta_drift": dd,
                    "prev_in_table": prev["in_table"],
                    "cur_in_table": p["in_table"],
                    "prev_type": prev["type"],
                    "cur_type": p["type"],
                    "prev_text": prev["text_preview"],
                    "cur_text": p["text_preview"],
                })
                print(f'{p["seq_idx"]:>5}  {prev["word_i"]:>6}->{p["word_i"]:<10}  '
                      f'{p["word_page"]:>3}  {p["oxi_page"]:>3}  '
                      f'{p["page_delta"]:>+4d}  {prev["drift"]:>+11.2f}  '
                      f'{p["drift"]:>+11.2f}  {dd:>+9.2f}  {p["text_preview"]}')
        prev = p
    print(f"  Total JUMP events: {len(jumps)}")

    # Cumulative drift trace every ~20 steps
    print()
    print("=== Cumulative drift trace (every ~5%) ===")
    print(f'{"seq":>5}  {"word_i":>5}  {"wp":>3}  {"op":>3}  {"Δp":>4}  {"drift_pt":>10}  type')
    step = max(1, len(valid_sorted) // 20)
    for i in range(0, len(valid_sorted), step):
        p = valid_sorted[i]
        print(f'{p["seq_idx"]:>5}  {p["word_i"]:>5}  {p["word_page"]:>3}  '
              f'{p["oxi_page"]:>3}  {p["page_delta"]:>+4d}  {p["drift"]:>+10.2f}  {p["type"]}')
    p = valid_sorted[-1]
    print(f'{p["seq_idx"]:>5}  {p["word_i"]:>5}  {p["word_page"]:>3}  '
          f'{p["oxi_page"]:>3}  {p["page_delta"]:>+4d}  {p["drift"]:>+10.2f}  {p["type"]}  (final)')

    # Aggregate per type
    type_summary = []
    by_type_records = defaultdict(list)
    for p in valid:
        by_type_records[p["type"]].append(p["drift"])
    for t, drifts in sorted(by_type_records.items(), key=lambda kv: -len(kv[1])):
        type_summary.append({
            "type": t,
            "count": len(drifts),
            "mean_abs_drift": round(statistics.fmean(drifts), 3),
            "median_abs_drift": round(statistics.median(drifts), 3),
        })

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    payload = {
        "doc_id": doc_id,
        "filename": fname,
        "page_geometry": geom,
        "n_matches": len(diff["matches"]),
        "n_valid": len(valid),
        "skipped": dict(skipped),
        "n_jumps": len(jumps),
        "jumps": jumps,
        "type_summary": type_summary,
        "paragraphs": paragraphs,
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {out_path}")


if __name__ == "__main__":
    main()
