"""Compute per-paragraph y-drift between Word and Oxi for a single doc.

For each Word paragraph, find the matching Oxi paragraph (by text-prefix
+ same page) and report `drift = word_y - oxi_y`. Sudden drift jumps
identify where Oxi's vertical layout diverges from Word's — typically
line-wrap count differences, skipped empty paragraphs, or header height
mismatch.

Output: pipeline_data/per_para_y_drift_<doc>.json

Usage:
    python tools/metrics/per_para_y_drift.py d77a58485f16
    python tools/metrics/per_para_y_drift.py 459f05f1e877
"""
from __future__ import annotations
import json
import os
import re
import sys

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
WORD_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_word")
OXI_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_oxi")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data")

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def normalize_text(s: str) -> str:
    if not s:
        return ""
    s = s.replace("　", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def compute_drift(doc_id: str) -> list[dict]:
    with open(os.path.join(WORD_DIR, f"{doc_id}.json"), encoding="utf-8") as f:
        w = json.load(f)
    with open(os.path.join(OXI_DIR, f"{doc_id}.json"), encoding="utf-8") as f:
        oxi = json.load(f)

    # Build Oxi index: page -> [(normalized_text, y, x, para_idx), ...]
    oxi_by_page: dict[int, list[tuple]] = {}
    for pg_str, entries in oxi["pages"].items():
        page = int(pg_str)
        for e in entries:
            t = normalize_text(e.get("text", ""))
            y = e.get("y") or 0
            x = e.get("x") or 0
            oxi_by_page.setdefault(page, []).append((t, y, x, e.get("para_idx")))

    # Track already-used Oxi entries to avoid double-matching
    used: set[tuple] = set()
    drift_records: list[dict] = []

    for wp in w["paragraphs"]:
        wt = normalize_text(wp.get("text", ""))
        if len(wt) < 3:
            continue
        wpage = wp.get("page")
        wy = wp.get("y")
        if wpage is None or wy is None:
            continue
        # Find Oxi match: prefix-equal, same page, not used
        best = None
        cand_list = oxi_by_page.get(wpage, [])
        for idx, (ot, oy, ox, opi) in enumerate(cand_list):
            if (wpage, idx) in used:
                continue
            n = min(len(wt), len(ot), 10)
            if n < 3:
                continue
            if wt[:n] == ot[:n]:
                # Prefer y close to expected (smaller |delta|)
                if best is None or abs(oy - wy) < abs(best[1] - wy):
                    best = (idx, oy, ox, opi)
        if best is None:
            continue
        idx, oy, ox, opi = best
        used.add((wpage, idx))
        drift = wp["y"] - oy
        drift_records.append({
            "w_i": wp["i"],
            "page": wpage,
            "word_y": round(wy, 1),
            "oxi_y": round(oy, 1),
            "drift": round(drift, 2),
            "text": wt[:50],
            "oxi_para_idx": opi,
        })
    return drift_records


def report(records: list[dict]) -> None:
    print(f'{"w_i":>4s} {"pg":>3s} {"w_y":>7s} {"o_y":>7s} {"drift":>8s} {"Δ":>6s} text')
    print("-" * 100)
    prev_drift = None
    for r in records:
        delta_str = ""
        if prev_drift is not None:
            d = r["drift"] - prev_drift
            if abs(d) >= 5:
                delta_str = f"{d:+.2f}"
        print(f'{r["w_i"]:>4d} {r["page"]:>3d} {r["word_y"]:>7.1f} {r["oxi_y"]:>7.1f} '
              f'{r["drift"]:>+8.2f} {delta_str:>6s} {r["text"][:50]!r}')
        prev_drift = r["drift"]


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: per_para_y_drift.py <doc_id_prefix>", file=sys.stderr)
        return 2
    prefix = sys.argv[1]
    # Find matching doc_id
    candidates = [f.removesuffix(".json") for f in os.listdir(WORD_DIR)
                  if f.startswith(prefix) and f.endswith(".json")]
    if not candidates:
        print(f"no doc found matching {prefix!r}", file=sys.stderr)
        return 1
    doc_id = candidates[0]
    print(f"Computing y-drift for {doc_id}...")
    records = compute_drift(doc_id)
    report(records)
    out_path = os.path.join(OUT_DIR, f"per_para_y_drift_{doc_id}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"doc_id": doc_id, "records": records}, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
