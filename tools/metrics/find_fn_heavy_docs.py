"""Find documents with footnotes in the baseline corpus."""
import os, sys, glob, time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = r"tools\golden-test\documents\docx"

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    results = []
    docx_paths = sorted(glob.glob(os.path.join(DOCX_DIR, "*.docx")))
    print(f"Scanning {len(docx_paths)} docs for footnotes...")
    for p in docx_paths:
        name = os.path.basename(p)
        try:
            doc = word.Documents.Open(os.path.abspath(p), ReadOnly=True)
            fn_count = doc.Footnotes.Count
            if fn_count > 0:
                # Count pages that have footnotes
                fn_pages = set()
                for fn in doc.Footnotes:
                    try:
                        fn_pages.add(fn.Reference.Information(3))
                    except:
                        pass
                total_pg = doc.ComputeStatistics(2)  # wdStatisticPages
                results.append({
                    "name": name,
                    "fn_count": fn_count,
                    "fn_pages": len(fn_pages),
                    "total_pages": total_pg,
                })
                print(f"  {name[:50]:50s} fn={fn_count:>3} pages_with_fn={len(fn_pages):>2} total={total_pg}")
            doc.Close(False)
        except Exception as e:
            print(f"  [ERR] {name}: {e}")
    # Sort by fn_count desc
    results.sort(key=lambda r: -r["fn_count"])
    print(f"\n=== TOP 10 fn-heavy docs ===")
    for r in results[:10]:
        print(f"  {r['name'][:50]:50s} fn={r['fn_count']:>3} pages_with_fn={r['fn_pages']}/{r['total_pages']}")
finally:
    word.Quit()
