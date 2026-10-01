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
# OXI_GDI_EXE lets a gate run against a SNAPSHOT of the renderer so a
# concurrent `cargo build --release` cannot swap the binary mid-run (the
# stale/partial-binary trap). Unset = the normal build output.
RENDERER = os.environ.get("OXI_GDI_EXE") or os.path.join(
    REPO_ROOT, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_oxi")


WORD_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_word")


def doc_id_from_filename(fname: str) -> str:
    """The id `pagination_diff` joins on -- i.e. the one the WORD side used.

    The historical rule is `base.split("_")[0]`, right for the corpus docs whose
    stem is `<hash>_<title>`. It COLLIDES for the handful named `<word>_<word>`:
    `gen2_079_Technical_Specification` -> `gen2`, `gen_tables` -> `gen`,
    `test_widow` -> `test`. Those three are stored on the Word side under their
    FULL basename, so the truncated file never matched and they dropped out of
    the gate silently -- the denominator read 93/93 instead of 96/96, while
    `pass_rate` stayed 100% either way (the n_total trap). Worse, every doc
    sharing a truncated prefix overwrote the same output file.

    Prefer the full basename whenever Word truth exists under it; otherwise keep
    the historical truncation, so every already-matching doc is untouched.
    """
    base = os.path.splitext(fname)[0]
    if os.path.exists(os.path.join(WORD_DIR, base + ".json")):
        return base
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
        # Group text elements by (para_idx, cell_para_idx).
        # R7.32 (Day 33 part 72, 2026-05-13): cell_para_idx distinguishes
        # paragraphs within the same table cell. Without it, all cell
        # paragraphs collapse under one para_idx (= table block_idx) and
        # the matcher misattributes diff matches (e3c545 # プレフィックス was
        # reported as +3 when actual layout delta was +1). Pre-R7.32 numbers
        # were inflated by hidden cell-paragraph deltas.
        groups: dict = {}
        for el in page.get("elements", []):
            if el.get("type") != "text":
                continue
            pi = el.get("para_idx")
            cpi = el.get("cell_para_idx")
            # R7.44 (Day 34 part 13, 2026-05-13): cell row/col indices added to
            # disambiguate cells that share (block_idx, cpi=0). Pre-R7.44 four
            # "千円" cells in one row collapsed into one "千円千円千円千円" record,
            # which forced the matcher into substring/multi-instance workarounds.
            cri = el.get("cell_row_idx")
            cci = el.get("cell_col_idx")
            if pi is None:
                # Pseudo-key: cluster by y-line (0.5pt) — matches cascade tool
                key = ("y", round(el["y"] * 2) / 2)
            else:
                # 2026-05-16 (Session 62, R7.77 post): nested-table disambiguation.
                # 3a4f9f p24 has 4 shift-variant nested tables under one outer
                # Block::Table (pi=186); cells from different variants share
                # (cri=0 cci=0) but appear at y=122, 232, 324, 416. Without
                # disambiguation, the matcher collapses them into one record with
                # capacity=1, producing false +3 deltas + 30 unmatched on 3a4f9f.
                # Conditional disambiguation: if an existing slot with the same
                # base key has y > 60pt away from current element, create a new
                # slot with an instance suffix. 60pt threshold is wider than
                # typical short-cell wrap (< 5 lines = 87pt) is OK for most cases
                # but disambiguates separated table instances (>60pt gap).
                base_key = (pi, cpi, cri, cci)
                # S1616: a NESTED table cell carries its ancestor path; without it
                # the inner cell (0,1) collided with the outer cell (0,1) and their
                # paragraphs merged (forms__005d851e).
                if el.get("cell_path"):
                    base_key = base_key + (tuple(tuple(x) for x in el["cell_path"]),)
                key = base_key
                instance = 0
                while key in groups and not el.get("vert") and (cpi is not None or cri is not None or cci is not None):
                    existing_y = groups[key]["y_min"]
                    if abs(el["y"] - existing_y) <= 60:
                        break  # close enough — same cell
                    instance += 1
                    key = base_key + (instance,)
            slot = groups.setdefault(key, {
                "para_idx": pi,
                # S177 (2026-05-22): preserve cell identity so the positional
                # matcher (pagination_diff_positional.py) can enumerate
                # paragraphs in document order. Text-prefix matching alone
                # produces false matches on docs with repeated short prefixes
                # (d4d126: 365 paras, many empty/□/○; misled S173 hypothesis).
                "cell_para_idx": cpi,
                "cell_row_idx": cri,
                "cell_col_idx": cci,
                "text_parts": [],
                "y_min": el["y"],
                "x_min": el["x"],
                # Session 74 Phase C (2026-05-17): track text_y_off of the
                # min-y line. After Phase D, y will be LINE BOX TOP directly;
                # for now Python tools subtract text_y_off to recover it. See
                # memory/session71_y_convention_refactor_design.md.
                "text_y_off_at_ymin": el.get("text_y_off", 0.0),
            })
            # Case conversion changes painted glyphs, not the source paragraph identity.
            text = el.get("source_text")
            if not isinstance(text, str):
                text = el.get("text", "")
            slot["text_parts"].append((el["y"], el["x"], text, bool(el.get("vert"))))
            if el["y"] < slot["y_min"]:
                slot["y_min"] = el["y"]
                slot["text_y_off_at_ymin"] = el.get("text_y_off", 0.0)
            slot["x_min"] = min(slot["x_min"], el["x"])
        # Build records
        records = []
        for key, slot in groups.items():
            # Sort by (y, x) so multi-line wrapped paragraphs concatenate
            # line-by-line top-to-bottom, not interleaved by X across lines.
            # Vertical elements are emitted in reading order: runs down each
            # column, columns right to left, then the next horizontal band.
            # Coordinate sorting loses that order when the first column is
            # indented or a column contains several run-sized elements.
            if (slot["para_idx"] is not None and slot["cell_para_idx"] is None and slot["cell_row_idx"] is None and slot["cell_col_idx"] is None) or (slot["text_parts"] and all(p[3] for p in slot["text_parts"])):
                pass
            else:
                slot["text_parts"].sort(key=lambda yxt: (yxt[0], yxt[1]))
            text = "".join(t for _, _, t, _ in slot["text_parts"])
            text = text.replace("\n", "").replace("\r", "")[:30]
            records.append({
                "para_idx": slot["para_idx"],
                "cell_para_idx": slot["cell_para_idx"],
                "cell_row_idx": slot["cell_row_idx"],
                "cell_col_idx": slot["cell_col_idx"],
                "text": text,
                "y": round(slot["y_min"], 2),
                "x": round(slot["x_min"], 2),
                "text_y_off": round(slot["text_y_off_at_ymin"], 2),
            })
        # Sort within page by Y, then X (reading order)
        records.sort(key=lambda r: (r["y"], r["x"]))
        out[str(page_num)] = records
    # 2026-09-26: a body paragraph that carries a page break INSIDE it
    # (`<w:t>    </w:t><w:br w:type="page"/><w:t>DESCRIPTION…</w:t>`,
    # educational__004cf6e9 pi=13) yields one whitespace-only record on the
    # page where it starts and one with its text on the next page. Word's
    # truth (Information(3) at the collapsed start) is the START page, and
    # the text-prefix matcher can only find the second record, so the doc
    # read as +1 although both engines start the paragraph on the same
    # page. Fold the continuation's text into the whitespace-only opener
    # and drop the continuation record.
    # The opener must be NON-EMPTY whitespace (the spaces before the break)
    # and the continuation must be the first record of the next page: para_idx
    # is a per-section block index, so an EMPTY paragraph ending one section's
    # last page shares its index with an unrelated paragraph of the next
    # section (technical__c6d7cdb1 / policies__0c94a7bc read -1 x 2..16 when
    # the empty opener was folded).
    pages_sorted = sorted(out.keys(), key=int)
    for a, b in zip(pages_sorted, pages_sorted[1:]):
        nxt = out[b]
        if not nxt:
            continue
        top_y = min(r["y"] for r in nxt)
        for rec in out[a]:
            if rec["para_idx"] is None or rec["cell_para_idx"] is not None or rec["cell_row_idx"] is not None:
                continue
            if not rec["text"] or rec["text"].replace("　", " ").strip():
                continue
            for j, cont in enumerate(nxt):
                if (cont["para_idx"] == rec["para_idx"] and cont["cell_para_idx"] is None and cont["cell_row_idx"] is None
                        and cont["text"].strip() and cont["y"] <= top_y + 0.01):
                    rec["text"] = (rec["text"] + cont["text"])[:30]
                    del nxt[j]
                    break
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
            timeout=300,
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


def measure_docs_batch(docx_paths, exe=None, env=None, stall_timeout=300):
    """Measure many documents with ONE renderer process (`--batch=LIST`).

    A renderer cold start costs ~1.25s of a ~1.6s median document, so this is
    what makes a 1076-document gate fast. Returns a list parallel to
    `docx_paths`: the measure_doc() dict, or an Exception. A document that makes
    no progress for `stall_timeout` seconds is killed and recorded as an error;
    the batch resumes with the next one.
    """
    import threading, queue
    exe = exe or RENDERER
    results = [None] * len(docx_paths)
    start = 0
    while start < len(docx_paths):
        with tempfile.TemporaryDirectory(prefix="oxi_batch_") as tmp:
            dumps = [os.path.join(tmp, f"{i}.json") for i in range(len(docx_paths))]
            lst = os.path.join(tmp, "list.txt")
            with open(lst, "w", encoding="utf-8") as f:
                for i in range(start, len(docx_paths)):
                    f.write(f"{docx_paths[i]}\t{dumps[i]}\n")
            proc = subprocess.Popen([exe, "--batch=" + lst], stdout=subprocess.PIPE,
                                    stderr=subprocess.DEVNULL, text=True, encoding="utf-8",
                                    env=env)
            q = queue.Queue()
            threading.Thread(target=lambda: [q.put(l) for l in proc.stdout] + [q.put(None)], daemon=True).start()
            i = start
            while i < len(docx_paths):
                try:
                    line = q.get(timeout=stall_timeout)
                except queue.Empty:
                    proc.kill()
                    results[i] = RuntimeError(f"stalled > {stall_timeout}s")
                    i += 1
                    break
                if line is None:  # process ended early (crash / memory cap)
                    results[i] = RuntimeError(f"renderer exited (rc={proc.wait()})")
                    i += 1
                    break
                if not line.startswith(("BATCH-OK ", "BATCH-ERR ")):
                    continue
                if line.startswith("BATCH-OK ") and os.path.exists(dumps[i]):
                    with open(dumps[i], encoding="utf-8") as f:
                        dump = json.load(f)
                    results[i] = {"filename": os.path.basename(docx_paths[i]),
                                  "n_pages": len(dump.get("pages", [])),
                                  "pages": aggregate_dump(dump)}
                else:
                    results[i] = RuntimeError("renderer failed on this document")
                i += 1
            proc.kill()
            proc.wait()
            start = i
    return results


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
