"""Phase A: COM ground truth measurement for holistic refactor.

For each baseline doc, measure Word's complete layout:
- Per-paragraph: page, text_top_y, x_start, line_count, line_ys (per-line top)
- Plus paragraph properties: indent_l, indent_r, first_line_indent
- Per-paragraph fn refs (footnote ids inline)
- Page setup: top/bottom margin, page width/height, body width

Output: pipeline_data/word_layout_truth/<doc_id>.json

Usage:
    python tools/metrics/measure_word_full_layout.py <doc_id_or_prefix>  # single doc
    python tools/metrics/measure_word_full_layout.py --all                # all baseline docs

Slow: ~20-60s per doc (depending on paragraph count, since each char's position
is sampled). Cached output means re-runs are fast (skip if exists).
"""
import json, os, sys, time
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
DOCS_DIR = REPO_ROOT / "tools" / "golden-test" / "documents" / "docx"
OUT_DIR = REPO_ROOT / "pipeline_data" / "word_layout_truth"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# Skip docs that are sandbox / generated:
SKIP_PREFIXES = {"gen", "gen2", "test", "pixel"}


def measure_doc(word, docx_path):
    """Run COM measurement on one docx. Returns dict for JSON output."""
    doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
    time.sleep(1.0)
    try:
        ps = doc.Sections(1).PageSetup
        page_w = ps.PageWidth
        page_h = ps.PageHeight
        top_m = ps.TopMargin
        bot_m = ps.BottomMargin
        left_m = ps.LeftMargin
        right_m = ps.RightMargin
        body_w = page_w - left_m - right_m

        n_paras = doc.Paragraphs.Count
        n_pages = None
        try:
            n_pages = int(doc.ComputeStatistics(2))  # wdStatisticPages
        except Exception:
            pass

        # Footnote map: id → (page, ref_y)
        try:
            n_fn = doc.Footnotes.Count
        except Exception:
            n_fn = 0
        fn_info = []
        for i in range(1, n_fn + 1):
            fn = doc.Footnotes(i)
            ref_rng = fn.Reference
            try:
                ref_pg = int(ref_rng.Information(3))
                ref_y = ref_rng.Information(6)
            except Exception:
                ref_pg, ref_y = None, None
            fn_info.append({"index": i, "ref_page": ref_pg, "ref_y": ref_y, "text_prefix": fn.Range.Text[:50]})

        sel = word.Selection
        para_records = []
        for pi in range(1, n_paras + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            raw_text = rng.Text or ""

            # Use collapsed start range (R30 fix) — Information() at active end
            # returns wrong page for paragraphs spanning page boundary
            visible_offset = 0
            for ch in raw_text:
                if ch in ("\x0c", "\x0b", "\r", "\n", "\x07"):
                    visible_offset += 1
                else:
                    break
            if visible_offset > 0 and visible_offset < len(raw_text):
                first_visible = rng.Start + visible_offset
            else:
                first_visible = rng.Start

            # Start position
            try:
                start_rng = doc.Range(first_visible, first_visible)
                start_pg = int(start_rng.Information(3))
                start_y = start_rng.Information(6)
                start_x = start_rng.Information(5)
            except Exception:
                start_pg, start_y, start_x = None, None, None

            # Walk chars to find ALL line tops (per-line y)
            line_ys = {}  # (pg, y_key) → first_x
            for ci in range(rng.Start, rng.End):
                sel.SetRange(ci, ci + 1)
                try:
                    y = sel.Information(6)
                    x = sel.Information(5)
                    pg = int(sel.Information(3))
                except Exception:
                    continue
                y_key = round(y * 2) / 2
                key = (pg, y_key)
                if key not in line_ys:
                    line_ys[key] = x

            # Sort lines by (page, y)
            sorted_keys = sorted(line_ys.keys())
            line_records = [
                {"page": k[0], "y": k[1], "x_start": line_ys[k]}
                for k in sorted_keys
            ]
            n_lines = len(line_records)

            # In table?
            try:
                in_table = bool(rng.Tables.Count)
            except Exception:
                in_table = False

            # Footnote refs in this paragraph (by text inspection: \x02 markers)
            fn_ref_count = raw_text.count("\x02")

            text = raw_text.replace("\r", "").replace("\x07", "").replace("\n", "").replace("\x0c", "").replace("\x0b", "")

            para_records.append({
                "i": pi,
                "start_page": start_pg,
                "start_y": round(start_y, 2) if start_y is not None else None,
                "start_x": round(start_x, 2) if start_x is not None else None,
                "n_lines": n_lines,
                "lines": line_records,
                "in_table": in_table,
                "fn_ref_count": fn_ref_count,
                "text_prefix": text[:80],
                "text_len": len(text),
            })

        return {
            "filename": os.path.basename(str(docx_path)),
            "page_w": page_w,
            "page_h": page_h,
            "top_margin": top_m,
            "bottom_margin": bot_m,
            "left_margin": left_m,
            "right_margin": right_m,
            "body_w": body_w,
            "n_pages": n_pages,
            "n_paragraphs": n_paras,
            "footnotes": fn_info,
            "paragraphs": para_records,
        }
    finally:
        doc.Close(False)


def doc_id_from_filename(fname):
    return fname.split("_")[0]


def main():
    args = sys.argv[1:]
    if not args or "--help" in args or "-h" in args:
        print(__doc__)
        return

    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False

    try:
        all_docs = sorted([f for f in DOCS_DIR.iterdir() if f.suffix == ".docx"])
        # baseline filter: not in SKIP_PREFIXES
        baseline = [f for f in all_docs if doc_id_from_filename(f.name)[:4] not in {p[:4] for p in SKIP_PREFIXES}]

        if "--all" in args:
            targets = baseline
        else:
            prefix = args[0]
            targets = [f for f in baseline if f.name.startswith(prefix)]
            if not targets:
                # try doc_id_from_filename
                targets = [f for f in baseline if doc_id_from_filename(f.name).startswith(prefix)]
        if not targets:
            print(f"No matching docs for {args[0]!r}")
            return

        print(f"Processing {len(targets)} docs...")
        for i, fp in enumerate(targets, 1):
            doc_id = doc_id_from_filename(fp.name)
            out_path = OUT_DIR / f"{doc_id}.json"
            if out_path.exists() and "--force" not in args:
                print(f"[{i}/{len(targets)}] SKIP (cached): {doc_id}")
                continue
            t0 = time.time()
            print(f"[{i}/{len(targets)}] {doc_id} ({fp.name})...", flush=True)
            try:
                result = measure_doc(word, fp)
                with open(out_path, "w", encoding="utf-8") as f:
                    json.dump(result, f, indent=1, ensure_ascii=False)
                elapsed = time.time() - t0
                print(f"   -> n_paras={result['n_paragraphs']} n_pages={result['n_pages']} ({elapsed:.1f}s)")
            except Exception as e:
                print(f"   ERROR: {e}")
    finally:
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
