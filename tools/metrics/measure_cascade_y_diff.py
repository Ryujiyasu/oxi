"""Measure per-paragraph Y position via Word COM for the bottom-5 floor docs.

Phase 1 Session 1 of cascade_unification_plan.md (commit 9e3638b).

Goal: produce per-paragraph (page, para_idx, text[:30], y_pt, x_pt, font,
size) records for each of the 5 bottom-5 floor docs. This is the Word
side; pair with extract_oxi_layout_y.py output to compute per-paragraph
delta and locate the largest per-element compression sources.

Output: pipeline_data/cascade_word_y/<doc_id>.json
        pipeline_data/cascade_word_y/_summary.json (count + page range per doc)

Run from repo root:
    python tools/metrics/measure_cascade_y_diff.py            # all 5 docs
    python tools/metrics/measure_cascade_y_diff.py 2ea81a     # smoke test
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
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "cascade_word_y")

# Bottom-5 floor docs (per cascade_unification_plan.md §1.1)
FLOOR_DOCS = [
    {
        "id": "d77a58",
        "filename": "d77a58485f16_20240705_resources_data_outline_08.docx",
        "floor_page": 7,
        "floor_ssim": 0.6268,
        "issue": "fn cascade",
    },
    {
        "id": "b837808",
        "filename": "b837808d0555_20240705_resources_data_guideline_02.docx",
        "floor_page": 6,
        "floor_ssim": 0.6449,
        "issue": "fn cascade",
    },
    {
        "id": "29dc6e",
        "filename": "29dc6e8943fe_order_01.docx",
        "floor_page": 5,
        "floor_ssim": 0.6636,
        "issue": "table cell complexity",
    },
    {
        "id": "2ea81a",
        "filename": "2ea81a8441cc_0025006-192.docx",
        "floor_page": 2,
        "floor_ssim": 0.6643,
        "issue": "form/page-break cascade",
    },
    {
        "id": "e3c545",
        "filename": "e3c545fac7a7_LOD_Handbook.docx",
        "floor_page": 11,
        "floor_ssim": 0.6649,
        "issue": "multi-page complex cascade",
    },
]


def measure_doc(word, info: dict) -> dict:
    docx_path = os.path.join(DOCS_DIR, info["filename"])
    if not os.path.exists(docx_path):
        raise FileNotFoundError(docx_path)

    doc = word.Documents.Open(docx_path, ReadOnly=True)
    time.sleep(0.4)
    try:
        ps = doc.PageSetup
        n_paras = doc.Paragraphs.Count
        page_setup = {
            "width_pt": ps.PageWidth,
            "height_pt": ps.PageHeight,
            "margin_left_pt": ps.LeftMargin,
            "margin_top_pt": ps.TopMargin,
            "margin_right_pt": ps.RightMargin,
            "margin_bottom_pt": ps.BottomMargin,
        }
        # ComputeStatistics(2) = wdStatisticPages (number of pages)
        try:
            n_pages = doc.ComputeStatistics(2)
        except Exception:
            n_pages = None

        rows = []
        for pi in range(1, n_paras + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            # CRITICAL (Round 30 fix): Range.Information returns the position of the
            # *active end* of the range. For paragraphs whose range spans a page
            # boundary (e.g., a paragraph at the end of a multi-page cell whose
            # trailing run/marker overflows to the next page), the end-position Y
            # is on the *next* page while the START is the actual visible top of
            # the paragraph. Round 28-29 used end-position Y and produced false
            # outliers for cell-end paragraphs (delta -361pt artifacts).
            # Use Range.Start collapsed to a zero-length range to get true start Y.
            try:
                start_rng = doc.Range(rng.Start, rng.Start)
                y = start_rng.Information(6)
                x = start_rng.Information(5)
                page = start_rng.Information(3)
                # Also capture end-position for diagnosis: nonzero start_vs_end
                # delta indicates a multi-page paragraph or a cell-trailing edge case.
                y_end = rng.Information(6)
                page_end = rng.Information(3)
            except Exception:
                y = x = page = y_end = page_end = None
            text = (rng.Text or "").replace("\r", "").replace("\x07", "").replace("\n", "")
            text = text[:30]

            font_name = ""
            font_size = None
            try:
                if rng.Runs.Count > 0:
                    r0 = rng.Runs(1)
                    font_name = r0.Font.Name
                    font_size = r0.Font.Size
            except Exception:
                pass

            try:
                pf = p.Format
                ls_rule = pf.LineSpacingRule
                ls = pf.LineSpacing
                sb = pf.SpaceBefore
                sa = pf.SpaceAfter
            except Exception:
                ls_rule = ls = sb = sa = None

            # Style name (helps map Heading/Normal/Body Text etc.)
            try:
                style_name = p.Style.NameLocal
            except Exception:
                style_name = ""

            # In-table flag (cell layout has different height calculation path)
            try:
                in_table = bool(rng.Tables.Count)
            except Exception:
                in_table = False

            rows.append({
                "i": pi,
                "page": page,
                "y_pt": y,
                "x_pt": x,
                "y_pt_end": y_end,
                "page_end": page_end,
                "text": text,
                "font": font_name,
                "size_pt": font_size,
                "ls_rule": ls_rule,
                "ls": ls,
                "sb": sb,
                "sa": sa,
                "style": style_name,
                "in_table": in_table,
            })

        return {
            "doc_id": info["id"],
            "filename": info["filename"],
            "floor_page": info["floor_page"],
            "floor_ssim": info["floor_ssim"],
            "issue": info["issue"],
            "page_setup": page_setup,
            "n_pages": n_pages,
            "n_paras": n_paras,
            "paragraphs": rows,
        }
    finally:
        doc.Close(SaveChanges=False)


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)

    target = sys.argv[1] if len(sys.argv) > 1 else None
    targets = [d for d in FLOOR_DOCS if target is None or d["id"].startswith(target)]
    if not targets:
        print(f"no doc matches prefix '{target}', expected one of: "
              f"{[d['id'] for d in FLOOR_DOCS]}", file=sys.stderr)
        return 2

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    summary = []
    try:
        for info in targets:
            print(f"=== {info['id']} | {info['filename']} ===")
            print(f"  floor: p.{info['floor_page']} SSIM={info['floor_ssim']} ({info['issue']})")
            t0 = time.time()
            try:
                result = measure_doc(word, info)
            except FileNotFoundError as e:
                print(f"  SKIP (missing): {e}")
                continue
            elapsed = time.time() - t0
            out_path = os.path.join(OUT_DIR, f"{info['id']}.json")
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            n_paras = result["n_paras"]
            n_pages = result["n_pages"]
            print(f"  -> {out_path}  ({n_paras} paras, {n_pages} pages, {elapsed:.1f}s)")
            summary.append({
                "doc_id": info["id"],
                "filename": info["filename"],
                "floor_page": info["floor_page"],
                "floor_ssim": info["floor_ssim"],
                "n_paras": n_paras,
                "n_pages": n_pages,
                "elapsed_sec": round(elapsed, 1),
            })
    finally:
        word.Quit()

    summary_path = os.path.join(OUT_DIR, "_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump({"docs": summary}, f, ensure_ascii=False, indent=2)
    print(f"summary -> {summary_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
