"""Drive oxi-gdi-renderer --dump-layout for each baseline doc, then
extract per-page (para_idx, text) records.

Phase 1 gate of the redesigned merge methodology (2026-04-28). Pair with
measure_pagination_word.py output via pagination_diff.py.

Renderer call (from main.rs:9-35):
    oxi-gdi-renderer.exe <input.docx> <output_prefix> [dpi] --dump-layout=<json>
The renderer returns early after dumping (no PNG generated).

Output: pipeline_data/pagination_oxi/<doc_id>.json (page → list of paragraph records)
        pipeline_data/pagination_oxi/_summary.json

Run from repo root:
    python tools/metrics/measure_pagination_oxi.py            # all docs
    python tools/metrics/measure_pagination_oxi.py 2ea81a     # prefix filter
    python tools/metrics/measure_pagination_oxi.py --limit=20

Pre-req: tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe must
exist and be up to date with the layout code under test. Build with
`cd tools/oxi-gdi-renderer && cargo build --release` first.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import time
from collections import defaultdict

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCS_DIR = os.path.join(REPO_ROOT, "tools", "golden-test", "documents", "docx")
RENDERER = os.path.join(REPO_ROOT, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_oxi")


def doc_id_from_filename(fname: str) -> str:
    base = os.path.splitext(fname)[0]
    return base.split("_")[0]


def aggregate_dump(dump: dict) -> dict:
    """Walk the renderer's JSON dump and build (page → [paragraph records]).

    A "paragraph record" here is (para_idx, text_prefix, x_min, y_min,
    is_in_table_guess). Text is the concatenation of text-element fragments
    sharing a para_idx within a single page (limited to first 30 chars).

    para_idx may be null for table-cell text in some renderer paths
    (matches behavior noted in cascade_cross_join.py). We retain those as
    pseudo-paragraphs keyed by (page, y-cluster) to avoid losing them.
    """
    out = {}
    for page in dump.get("pages", []):
        page_num = page["page"]
        # Group text elements by para_idx (preserving insertion order via dict)
        groups: dict = {}
        for el in page.get("elements", []):
            if el.get("type") != "text":
                continue
            key = el.get("para_idx")
            if key is None:
                # Pseudo-key: cluster by y-line (0.5pt) — matches cascade tool
                key = ("y", round(el["y"] * 2) / 2)
            slot = groups.setdefault(key, {
                "para_idx": el.get("para_idx"),
                "text_parts": [],
                "y_min": el["y"],
                "x_min": el["x"],
            })
            slot["text_parts"].append((el["x"], el.get("text", "")))
            slot["y_min"] = min(slot["y_min"], el["y"])
            slot["x_min"] = min(slot["x_min"], el["x"])
        # Build records
        records = []
        for key, slot in groups.items():
            slot["text_parts"].sort(key=lambda xt: xt[0])
            text = "".join(t for _, t in slot["text_parts"])
            text = text.replace("\n", "").replace("\r", "")[:30]
            records.append({
                "para_idx": slot["para_idx"],
                "text": text,
                "y": round(slot["y_min"], 2),
                "x": round(slot["x_min"], 2),
            })
        # Sort within page by Y, then X (reading order)
        records.sort(key=lambda r: (r["y"], r["x"]))
        out[str(page_num)] = records
    return out


def measure_doc(docx_path: str) -> dict:
    with tempfile.TemporaryDirectory(prefix="oxi_dump_") as tmp:
        out_prefix = os.path.join(tmp, "page_")
        dump_path = os.path.join(tmp, "layout.json")
        # Renderer requires output_prefix even when dumping (positional arg).
        proc = subprocess.run(
            [RENDERER, docx_path, out_prefix, "--dump-layout=" + dump_path],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if proc.returncode != 0:
            raise RuntimeError(f"renderer failed (rc={proc.returncode}): {proc.stderr[:500]}")
        if not os.path.exists(dump_path):
            raise RuntimeError(f"dump not produced (stderr: {proc.stderr[:500]})")
        with open(dump_path, encoding="utf-8") as f:
            dump = json.load(f)

    by_page = aggregate_dump(dump)
    return {
        "filename": os.path.basename(docx_path),
        "n_pages": len(dump.get("pages", [])),
        "pages": by_page,
    }


def main() -> int:
    if not os.path.exists(RENDERER):
        print(f"renderer not found at {RENDERER}", file=sys.stderr)
        print("Build it first: cd tools/oxi-gdi-renderer && cargo build --release", file=sys.stderr)
        return 2
    os.makedirs(OUT_DIR, exist_ok=True)

    prefix = None
    limit = None
    for arg in sys.argv[1:]:
        if arg.startswith("--limit="):
            limit = int(arg.split("=", 1)[1])
        elif not arg.startswith("--"):
            prefix = arg

    all_docx = sorted(
        f for f in os.listdir(DOCS_DIR)
        if f.lower().endswith(".docx") and not f.startswith("~$")
    )
    if prefix:
        all_docx = [f for f in all_docx if f.startswith(prefix) or doc_id_from_filename(f).startswith(prefix)]
    if limit:
        all_docx = all_docx[:limit]
    if not all_docx:
        print(f"no docx matched (prefix={prefix})", file=sys.stderr)
        return 2

    summary = []
    for fname in all_docx:
        docx_path = os.path.join(DOCS_DIR, fname)
        doc_id = doc_id_from_filename(fname)
        print(f"=== {doc_id} | {fname} ===")
        t0 = time.time()
        try:
            result = measure_doc(docx_path)
        except Exception as e:
            print(f"  FAIL: {e}", file=sys.stderr)
            summary.append({"doc_id": doc_id, "filename": fname, "error": str(e)[:200]})
            continue
        elapsed = time.time() - t0
        out_path = os.path.join(OUT_DIR, f"{doc_id}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"  -> {out_path}  ({result['n_pages']} pages, {elapsed:.1f}s)")
        summary.append({
            "doc_id": doc_id,
            "filename": fname,
            "n_pages": result["n_pages"],
            "elapsed_sec": round(elapsed, 1),
        })

    summary_path = os.path.join(OUT_DIR, "_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump({"docs": summary}, f, ensure_ascii=False, indent=2)
    print(f"summary -> {summary_path}  ({len(summary)} docs)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
