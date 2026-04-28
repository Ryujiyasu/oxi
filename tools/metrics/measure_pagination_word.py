"""Measure per-paragraph (page, text) via Word COM for the baseline docs.

Phase 1 gate of the redesigned merge methodology (2026-04-28). Pagination
correctness is the primary signal: per doc, compare Word page N's
paragraph set against Oxi page N's set. See memory file
methodology_phase_based.md for the strategic context.

Lighter than measure_cascade_y_diff.py — only need (i, page, text[:30])
per paragraph, not Y/X/font/spacing. Same R30 fix (collapsed start range)
to avoid Information(3) reporting end-of-range page on multi-page paragraphs.

Output: pipeline_data/pagination_word/<doc_id>.json
        pipeline_data/pagination_word/_summary.json

Run from repo root:
    python tools/metrics/measure_pagination_word.py            # all docs in DOCS_DIR
    python tools/metrics/measure_pagination_word.py 2ea81a     # prefix filter (one doc)
    python tools/metrics/measure_pagination_word.py --limit=20 # first 20 docs
"""
from __future__ import annotations

import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCS_DIR = os.path.join(REPO_ROOT, "tools", "golden-test", "documents", "docx")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_word")


def doc_id_from_filename(fname: str) -> str:
    """Match cascade tools' convention: first underscore-separated token."""
    base = os.path.splitext(fname)[0]
    return base.split("_")[0]


def measure_doc(word, docx_path: str) -> dict:
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    time.sleep(0.3)
    try:
        try:
            n_pages = doc.ComputeStatistics(2)  # wdStatisticPages
        except Exception:
            n_pages = None
        n_paras = doc.Paragraphs.Count

        rows = []
        for pi in range(1, n_paras + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            try:
                # R30 fix: collapsed start range avoids Information() reporting
                # the active-end page for paragraphs whose trailing run/marker
                # overflows to the next page.
                start_rng = doc.Range(rng.Start, rng.Start)
                page = start_rng.Information(3)  # wdActiveEndPageNumber
            except Exception:
                page = None
            text = (rng.Text or "").replace("\r", "").replace("\x07", "").replace("\n", "")
            text = text[:30]

            try:
                in_table = bool(rng.Tables.Count)
            except Exception:
                in_table = False

            rows.append({
                "i": pi,
                "page": page,
                "text": text,
                "in_table": in_table,
            })

        return {
            "filename": os.path.basename(docx_path),
            "n_pages": n_pages,
            "n_paras": n_paras,
            "paragraphs": rows,
        }
    finally:
        doc.Close(SaveChanges=False)


def main() -> int:
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
        print(f"no docx matched (prefix={prefix}, dir={DOCS_DIR})", file=sys.stderr)
        return 2

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    summary = []
    try:
        for fname in all_docx:
            docx_path = os.path.join(DOCS_DIR, fname)
            doc_id = doc_id_from_filename(fname)
            print(f"=== {doc_id} | {fname} ===")
            t0 = time.time()
            try:
                result = measure_doc(word, docx_path)
            except Exception as e:
                print(f"  FAIL: {e}", file=sys.stderr)
                summary.append({"doc_id": doc_id, "filename": fname, "error": str(e)})
                continue
            elapsed = time.time() - t0
            out_path = os.path.join(OUT_DIR, f"{doc_id}.json")
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"  -> {out_path}  ({result['n_paras']} paras, {result['n_pages']} pages, {elapsed:.1f}s)")
            summary.append({
                "doc_id": doc_id,
                "filename": fname,
                "n_paras": result["n_paras"],
                "n_pages": result["n_pages"],
                "elapsed_sec": round(elapsed, 1),
            })
    finally:
        word.Quit()

    summary_path = os.path.join(OUT_DIR, "_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump({"docs": summary}, f, ensure_ascii=False, indent=2)
    print(f"summary -> {summary_path}  ({len(summary)} docs)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
