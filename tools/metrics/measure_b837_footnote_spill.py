"""Measure b837 footnote distribution: fn# per page + spill detection.

Key questions per Task #3:
- Does any footnote body SPILL across pages (same fn# rendered on p.N and p.N+1)?
- Where are fn references located on body pages? Correlate reference page
  with body page.
- Confirm or refute oxi-2's "Word splits long fn bodies across pages"
  hypothesis.
"""
import os, sys, time, subprocess, json
from pathlib import Path
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = str(Path(r"tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx").resolve())

WD_FOOTNOTES_STORY = 2
WD_VPOS_RELATIVE_TO_PAGE = 6
WD_PAGE_NUMBER = 3


def measure():
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, timeout=5)
    time.sleep(0.5)
    word = win32com.client.DispatchEx("Word.Application")
    try:
        try: word.Visible = False
        except: pass
        try: word.DisplayAlerts = False
        except: pass
        doc = word.Documents.Open(DOC, ReadOnly=True); time.sleep(0.5)
        doc.Repaginate()

        result = {"footnote_refs": [], "footnote_bodies": []}

        # 1) For each footnote, find its reference location and body location
        fn_count = doc.Footnotes.Count
        print(f"Total footnotes: {fn_count}", file=sys.stderr)
        for fn_i in range(1, fn_count + 1):
            try:
                fn = doc.Footnotes(fn_i)
                # Reference position (the superscript number in body)
                ref = fn.Reference
                ref_page = ref.Information(WD_PAGE_NUMBER)
                ref_y = ref.Information(WD_VPOS_RELATIVE_TO_PAGE)
                # Body — first and last character positions
                body_range = fn.Range
                body_first_y = body_range.Information(WD_VPOS_RELATIVE_TO_PAGE)
                body_first_page = body_range.Information(WD_PAGE_NUMBER)
                # Try last char
                try:
                    last_ch = body_range.Characters(body_range.Characters.Count)
                    body_last_y = last_ch.Information(WD_VPOS_RELATIVE_TO_PAGE)
                    body_last_page = last_ch.Information(WD_PAGE_NUMBER)
                except Exception:
                    body_last_y = None
                    body_last_page = None
                text_sample = body_range.Text[:40].replace("\r", "").replace("\x07", "")
                entry = {
                    "fn_num": fn_i,
                    "ref_page": int(ref_page),
                    "ref_y": round(ref_y, 2),
                    "body_first_page": int(body_first_page),
                    "body_first_y": round(body_first_y, 2),
                    "body_last_page": int(body_last_page) if body_last_page else None,
                    "body_last_y": round(body_last_y, 2) if body_last_y else None,
                    "spill": (body_last_page != body_first_page) if body_last_page else None,
                    "text_sample": text_sample,
                }
                result["footnote_refs"].append(entry)
                print(f"fn{fn_i}: ref p{ref_page} y={ref_y:.1f} | body p{body_first_page} y={body_first_y:.1f}..p{body_last_page} y={body_last_y}", file=sys.stderr)
            except Exception as e:
                print(f"fn{fn_i}: err {e}", file=sys.stderr)
        doc.Close(False)
        return result
    finally:
        try: word.Quit()
        except: pass


def main():
    r = measure()
    out = "pipeline_data/b837_footnote_spill.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(r, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {out}")

    # Analysis summary
    fns = r["footnote_refs"]
    spills = [f for f in fns if f.get("spill")]
    print(f"\n=== Summary ===")
    print(f"Total footnotes: {len(fns)}")
    print(f"Footnotes with body spill across pages: {len(spills)}")
    if spills:
        for s in spills:
            print(f"  fn{s['fn_num']}: ref p{s['ref_page']} body p{s['body_first_page']}→p{s['body_last_page']}")
    else:
        print("  → NO footnote bodies spill across pages")

    # Per-page fn count
    from collections import Counter
    ref_pages = Counter(f["ref_page"] for f in fns)
    body_pages = Counter(f["body_first_page"] for f in fns)
    print(f"\nRef-page distribution: {dict(ref_pages)}")
    print(f"Body-page distribution: {dict(body_pages)}")
    # Mismatches: fn whose body first-page != ref page
    mismatches = [f for f in fns if f["ref_page"] != f["body_first_page"]]
    print(f"\nFn refs NOT on same page as body: {len(mismatches)}")
    for m in mismatches[:8]:
        print(f"  fn{m['fn_num']}: ref p{m['ref_page']} vs body p{m['body_first_page']}")


if __name__ == "__main__":
    main()
