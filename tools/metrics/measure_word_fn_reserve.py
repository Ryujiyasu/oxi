"""Multi-doc COM measurement of Word's fn reserve algorithm.

For each specified doc, for each page with footnotes:
  - body_bot: y of last body para + estimated line_h
  - fn_area_top: y of first fn body on that page
  - gap: fn_area_top - body_bot
  - fn_count: how many fns ref'd on that page
  - last_body_para_y: y of the final body paragraph
  - last_body_para_refs: fn refs inside that paragraph

Hypothesis check:
  - If Word uses per-line streaming reserve, last body para having MANY refs
    should STILL fit (not widow'd) because refs reserve as they go.
  - If Word pre-reserves all refs, last body para with many refs should be
    pushed to next page (like Oxi does).

Output: pipeline_data/ra_manual_measurements.json entry.
"""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Docs to measure. b837 already measured; augment with 2 more fn-heavy docs.
# Will be filled in based on find_fn_heavy_docs.py scan result.
DOCS = [
    r"tools\golden-test\documents\docx\b837808d0555_20240705_resources_data_guideline_02.docx",
    # Placeholder — fill after scan:
    # r"tools\golden-test\documents\docx\<doc2>.docx",
    # r"tools\golden-test\documents\docx\<doc3>.docx",
]

LINE_H_DEFAULT = 18.0  # pt for ~10.5pt font


def measure_doc(word, docx_path):
    """Return list of per-page measurements for this doc."""
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    time.sleep(0.5)
    try:
        result = {
            "doc": os.path.basename(docx_path),
            "total_pages": int(doc.ComputeStatistics(2)),
            "fn_total": doc.Footnotes.Count,
            "pages": [],
        }

        # Collect per-page body paragraphs
        by_page_paras = {}
        for i, p in enumerate(doc.Paragraphs, 1):
            try:
                pg = p.Range.Information(3)
                y = p.Range.Information(6)
                refs = []
                for r in p.Range.Footnotes:
                    refs.append(r.Index)
                by_page_paras.setdefault(pg, []).append({
                    "idx": i, "y": round(y, 2), "refs": refs,
                    "text": p.Range.Text[:40].replace('\r','').replace('\n',' ').replace('\x07',''),
                })
            except Exception:
                pass

        # Collect per-page fn body positions
        by_page_fns = {}
        for fn in doc.Footnotes:
            try:
                ref_pg = fn.Reference.Information(3)
                body_pg = fn.Range.Information(3)
                body_y = fn.Range.Information(6)
                ref_y = fn.Reference.Information(6)
                by_page_fns.setdefault(body_pg, []).append({
                    "seq": fn.Index, "ref_pg": ref_pg, "ref_y": round(ref_y, 2),
                    "body_y": round(body_y, 2),
                })
            except Exception:
                pass

        # Summarize per page
        for pg in sorted(by_page_paras.keys()):
            paras = by_page_paras[pg]
            fns = by_page_fns.get(pg, [])
            if not fns:
                continue
            # Filter paras actually on this page (idx may belong via para-split to another)
            paras_on_pg = [p for p in paras if 50 < p["y"] < 800]  # within body area
            paras_on_pg.sort(key=lambda x: x["y"])
            last_body = paras_on_pg[-1] if paras_on_pg else None
            fn_sorted = sorted(fns, key=lambda x: x["body_y"])
            fn_first = fn_sorted[0] if fn_sorted else None
            fn_last = fn_sorted[-1] if fn_sorted else None
            page_info = {
                "page": pg,
                "body_paras_count": len(paras_on_pg),
                "fn_count": len(fns),
                "last_body_para": last_body,
                "body_bot_approx": round(last_body["y"] + LINE_H_DEFAULT, 2) if last_body else None,
                "fn_area_top": fn_first["body_y"] if fn_first else None,
                "fn_area_bot": fn_last["body_y"] if fn_last else None,
                "fn_area_content_h": round(fn_last["body_y"] - fn_first["body_y"], 2) if fn_first and fn_last else None,
                "gap_body_to_fn": round(
                    (fn_first["body_y"] - (last_body["y"] + LINE_H_DEFAULT)), 2
                ) if last_body and fn_first else None,
                "fns_detail": fn_sorted,
            }
            result["pages"].append(page_info)
        return result
    finally:
        doc.Close(False)


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        all_results = []
        for d in DOCS:
            try:
                print(f"--- {os.path.basename(d)} ---")
                r = measure_doc(word, d)
                all_results.append(r)
                for pg in r["pages"]:
                    print(
                        f"  p{pg['page']}: body_bot~{pg['body_bot_approx']} "
                        f"fn_top={pg['fn_area_top']} gap={pg['gap_body_to_fn']} "
                        f"fn_count={pg['fn_count']} "
                        f"last_body_refs={pg['last_body_para']['refs'] if pg['last_body_para'] else '?'}"
                    )
            except Exception as e:
                print(f"  [ERR] {e}")
        # Save
        out = r"pipeline_data\word_fn_reserve_measurements.json"
        with open(out, "w", encoding="utf-8") as f:
            json.dump(all_results, f, ensure_ascii=False, indent=2)
        print(f"\nSaved {out}")
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
