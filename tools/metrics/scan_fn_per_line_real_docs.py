"""
Ra: scan baseline docs for real-world evidence of per-line fn commit.

Pattern sought: a paragraph that straddles pages AND has a fn-ref whose
anchor lands on a line after the first line. Target: confirm 2+ real docs
showing this pattern (to satisfy 3-doc + minimal-repro rule).

For each candidate doc, iterate paragraphs. For any paragraph with >=1
footnote reference, inspect:
  - Paragraph line placement (approximate via char-by-char Information(3,6)).
  - Which lines contain footnote refs (via footnote Reference anchor).
  - Does the paragraph straddle pages (first-char page != last-char page)?
  - Is the fn's anchor on a line that's on a DIFFERENT page than line 1?

Only do deep scan on docs where Footnotes.Count > 0.
"""
import win32com.client, json, os, glob, sys

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

DOCX_DIR = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx"
candidates = sorted(glob.glob(os.path.join(DOCX_DIR, "*.docx")))
# Skip temp files.
candidates = [c for c in candidates if "_pixel_tmp" not in c.lower()]

results = []
hits = []

print(f"Scanning {len(candidates)} documents for fn-ref per-line-commit evidence")

for idx, path in enumerate(candidates):
    doc_name = os.path.splitext(os.path.basename(path))[0]
    try:
        wdoc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
    except Exception as e:
        print(f"[{idx}/{len(candidates)}] OPEN_FAIL {doc_name}: {e}")
        continue
    try:
        fn_count = wdoc.Footnotes.Count
        if fn_count == 0:
            wdoc.Close(False)
            continue
        wdoc.Repaginate()
        # Build fn anchor Index → (page, char_pos, para_idx)
        fn_anchors = []
        for fi in range(1, fn_count + 1):
            fn = wdoc.Footnotes(fi)
            anchor = fn.Reference  # Range where the footnote's ref lives
            try:
                a_page = anchor.Information(3)
                a_y = anchor.Information(6)
                a_start = anchor.Start
                fn_page = fn.Range.Information(3)
            except Exception:
                continue
            fn_anchors.append({
                "idx": fi,
                "anchor_page": a_page,
                "anchor_y": round(a_y, 2),
                "anchor_start": a_start,
                "fn_page": fn_page,
            })
        if not fn_anchors:
            wdoc.Close(False)
            continue

        # For each paragraph with a fn_ref, check if it straddles pages and
        # which line the fn's anchor lives on.
        doc_hits = []
        for pi in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(pi)
            r = para.Range
            p_start = r.Start
            p_end = r.End
            # Any fn anchor within this paragraph?
            contained = [a for a in fn_anchors if p_start <= a["anchor_start"] < p_end]
            if not contained:
                continue
            # First-char and last-char pages.
            try:
                first_page = wdoc.Range(p_start, p_start + 1).Information(3)
                last_page = wdoc.Range(max(p_start, p_end - 2), p_end - 1).Information(3)
            except Exception:
                continue
            straddles = first_page != last_page
            if not straddles:
                continue
            # Find first-line y for this paragraph.
            try:
                first_y = round(wdoc.Range(p_start, p_start + 1).Information(6), 2)
            except Exception:
                first_y = None
            # For each fn anchor, line it's on is determined by its y.
            # Cluster anchors by (page, y_rounded). If anchor's (page, y_round) !=
            # (first_page, first_y_round), the anchor is on a later line/page.
            late_anchors = []
            for a in contained:
                y = a["anchor_y"]
                # Compare to first-line y. Different y means different line.
                # "Later line" = (anchor_page > first_page) OR
                # (same page & anchor_y > first_y + 1pt).
                if a["anchor_page"] > first_page:
                    late_anchors.append(a)
                elif first_y is not None and a["anchor_page"] == first_page and y > first_y + 1.0:
                    late_anchors.append(a)
            if late_anchors:
                # This paragraph has a fn ref on a LATER line.
                # Capture details.
                doc_hits.append({
                    "doc": doc_name,
                    "para_idx": pi,
                    "para_text_head": r.Text.strip()[:60],
                    "first_page": first_page,
                    "last_page": last_page,
                    "first_y": first_y,
                    "late_anchors": late_anchors,
                    "all_anchors_in_para": contained,
                })
        if doc_hits:
            print(f"[{idx}/{len(candidates)}] {doc_name}: fn={fn_count}  hits={len(doc_hits)}")
            for h in doc_hits:
                print(f"    para#{h['para_idx']} pages {h['first_page']}→{h['last_page']}  "
                      f"first_y={h['first_y']}  late_anchors={len(h['late_anchors'])}")
                for la in h["late_anchors"]:
                    print(f"      fn#{la['idx']} anchor p{la['anchor_page']} y={la['anchor_y']} "
                          f"fn_on_page={la['fn_page']}")
            hits.extend(doc_hits)
        else:
            # Still record that this doc has fn
            pass
    except Exception as e:
        print(f"[{idx}/{len(candidates)}] ERR {doc_name}: {e}")
    finally:
        try:
            wdoc.Close(False)
        except Exception:
            pass

    # Stop early if we have 3 distinct docs with hits.
    distinct = len({h["doc"] for h in hits})
    if distinct >= 3:
        print(f"\n  Found {distinct} distinct docs with hits — stopping early.")
        break

word.Quit()

out = os.path.join(os.path.dirname(__file__), 'output', 'scan_fn_per_line_real_docs.json')
os.makedirs(os.path.dirname(out), exist_ok=True)
with open(out, 'w', encoding='utf-8') as f:
    json.dump(hits, f, indent=2, ensure_ascii=False)

print(f"\n========== SUMMARY ==========")
print(f"Docs with hits: {len({h['doc'] for h in hits})}")
print(f"Total hits: {len(hits)}")
print(f"Saved: {out}")
for h in hits[:8]:
    print(f"  {h['doc']} para#{h['para_idx']}  p{h['first_page']}→p{h['last_page']}  "
          f"late={len(h['late_anchors'])}  [{h['para_text_head']}]")
