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
    # S432: some baseline docs (3a4f) fail/hang on Documents.Open of the
    # original (protected-view / lock); a %TEMP% copy opens cleanly (per
    # memory session424). Always open a temp copy — cheap and robust.
    # Use a UNIQUE temp filename (mkstemp) + cleanup: a fixed name collided
    # with a still-open/locked copy from a prior run (Permission denied),
    # silently failing every doc.
    import shutil, tempfile
    fd, tmp_copy = tempfile.mkstemp(suffix=".docx", prefix="ppw_")
    os.close(fd)
    shutil.copy(docx_path, tmp_copy)
    doc = word.Documents.Open(tmp_copy, ReadOnly=True)
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
            raw_text = rng.Text or ""
            try:
                # R30 fix: collapsed start range avoids Information() reporting
                # the active-end page for paragraphs whose trailing run/marker
                # overflows to the next page.
                # R39 (2026-04-28): skip leading control chars (page break
                # \x0c, line break \x0b, paragraph mark \r) when determining
                # the visible page. Paragraphs starting with <w:br type="page"/>
                # have rng.Start at the page-break marker (page N-1) but the
                # visible text on page N. Using start position misreports the
                # page in pagination_diff (0e7af1ae i=59 "総則" was a false
                # FAIL: visible text on page 2, marker on page 1).
                visible_offset = 0
                for ch in raw_text:
                    if ch in ("\x0c", "\x0b", "\r", "\n", "\x07"):
                        visible_offset += 1
                    else:
                        break
                if visible_offset > 0 and visible_offset < len(raw_text):
                    vis_rng = doc.Range(rng.Start + visible_offset,
                                        rng.Start + visible_offset)
                    page = vis_rng.Information(3)
                    y = vis_rng.Information(6)  # wdVerticalPositionRelativeToPage
                    x = vis_rng.Information(5)  # wdHorizontalPositionRelativeToPage
                else:
                    start_rng = doc.Range(rng.Start, rng.Start)
                    page = start_rng.Information(3)  # wdActiveEndPageNumber
                    y = start_rng.Information(6)
                    x = start_rng.Information(5)
                # S722 (2026-07-03): SHORT-PARAGRAPH page-boundary quirk. For a
                # 1-line paragraph pushed to the next page (tokyoshugyo wi=374
                # （セクシュアルハラスメントの禁止）, 17 visible chars: the
                # SCREEN render, the PDF export and Oxi all place it at p18
                # TOP), the collapsed-START query keeps answering the PRE-BREAK
                # logical anchor (p17, y=748.2 — reproducible even after
                # doc.Repaginate(); (s,s+1) answers p17 too), while the
                # active-END query answers the rendered page (p18, end-char
                # y=102.7 = the p18 top line). A SHORT body paragraph (≤20
                # visible chars at body width) cannot truly span two pages, so
                # start-page+1 == end-page on one is the API quirk -> take the
                # END page (and its y/x). Multi-line paragraphs keep the R30
                # collapsed-start (their end legitimately falls on later
                # pages). NOTE ComputeStatistics(wdStatisticLines) is NOT
                # usable as the 1-line test — it answers 3 for this visibly
                # 1-line heading. Table paragraphs are excluded (narrow cell
                # columns can wrap a short paragraph across a real page break).
                try:
                    end_page = rng.Information(3)  # active end
                    vis_chars = sum(1 for c in raw_text
                                    if c not in ("\x0c", "\x0b", "\r", "\n", "\x07"))
                    # Table membership via the START-collapsed range's
                    # Information(12): rng.Tables.Count is 1 for a BODY
                    # paragraph whose ¶ merely TOUCHES a following table
                    # (the wi=374 heading precedes the 第１３条 box).
                    in_tbl_start = bool(
                        doc.Range(rng.Start, rng.Start).Information(12))
                    if (end_page and page and end_page == page + 1
                            and vis_chars <= 20
                            and not in_tbl_start):
                        epos = max(rng.Start, rng.End - 1)
                        end_rng = doc.Range(epos, epos)
                        page = end_page
                        y = end_rng.Information(6)
                        x = end_rng.Information(5)
                except Exception:
                    pass
            except Exception:
                page = None
                y = None
                x = None
            text = raw_text.replace("\r", "").replace("\x07", "").replace("\n", "").replace("\x0c", "").replace("\x0b", "")
            text = text[:30]

            try:
                in_table = bool(rng.Tables.Count)
            except Exception:
                in_table = False

            # S432 (2026-05-29): structural cell coordinates for cell-aware
            # diagnostics (cell_iou_diff.py). The element_iou next-y height
            # derivation breaks inside tables because COM enumerates cells
            # row-major but Information(6) y is non-monotonic across a row
            # (S431 tokumei -81.6 artifact). With (table_start, row, col) we
            # can derive per-column-monotonic heights instead. Additive
            # fields only — existing pagination_diff / element_iou ignore them.
            cell_row = None      # 0-based row within outermost table
            cell_col = None      # 0-based column within outermost table
            table_start = None   # outermost-table Range.Start = stable table id
            if in_table:
                try:
                    cell = rng.Cells(1)            # innermost cell
                    cell_row = cell.RowIndex - 1
                    cell_col = cell.ColumnIndex - 1
                    # Outermost table = the one with the smallest Start among
                    # the tables overlapping this range (encloses the rest).
                    tcount = rng.Tables.Count
                    starts = []
                    for ti in range(1, tcount + 1):
                        try:
                            starts.append(rng.Tables(ti).Range.Start)
                        except Exception:
                            pass
                    if starts:
                        table_start = min(starts)
                except Exception:
                    # Empty end-of-cell / boundary paragraphs raise; leave None.
                    pass

            rows.append({
                "i": pi,
                "page": page,
                "y": round(y, 2) if y is not None else None,
                "x": round(x, 2) if x is not None else None,
                "text": text,
                "in_table": in_table,
                "cell_row": cell_row,
                "cell_col": cell_col,
                "table_start": table_start,
            })

        return {
            "filename": os.path.basename(docx_path),
            "n_pages": n_pages,
            "n_paras": n_paras,
            "paragraphs": rows,
        }
    finally:
        doc.Close(SaveChanges=False)
        try:
            os.remove(tmp_copy)
        except Exception:
            pass


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)

    prefix = None
    limit = None
    resume = False
    for arg in sys.argv[1:]:
        if arg.startswith("--limit="):
            limit = int(arg.split("=", 1)[1])
        elif arg == "--resume":
            # S432: skip docs whose json already carries S432 cell coords
            # (any paragraph with a `cell_row` key). Lets a killed full run
            # resume without redoing completed docs.
            resume = True
        elif not arg.startswith("--"):
            prefix = arg

    all_docx = sorted(
        f for f in os.listdir(DOCS_DIR)
        if f.lower().endswith(".docx") and not f.startswith("~$")
    )
    if prefix:
        all_docx = [f for f in all_docx if f.startswith(prefix) or doc_id_from_filename(f).startswith(prefix)]
    if resume:
        def _has_coords(fn):
            p = os.path.join(OUT_DIR, f"{doc_id_from_filename(fn)}.json")
            if not os.path.exists(p):
                return False
            try:
                with open(p, encoding="utf-8") as fh:
                    d = json.load(fh)
                return any("cell_row" in r for r in d.get("paragraphs", []))
            except Exception:
                return False
        all_docx = [f for f in all_docx if not _has_coords(f)]
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
