"""COM-scan Bug B family docs for paragraphs that have BOTH firstLineChars
and firstLine in XML, then measure Word's actual rendered x_start.

Goal: identify the trigger condition for Word's "ignore firstLine" behavior
observed in bd90b00 para 24 (Word renders at page margin x=56.5 despite
firstLineChars=100 firstLine=178 in XML).

For each candidate paragraph:
- Extract XML indent properties (left, firstLineChars, firstLine, leftChars)
- Extract paragraph leading whitespace (count of leading space chars + their font size)
- COM-measure Word's first-char x position
- Compute expected_x = page_margin_left + left_indent + firstLine_indent
- Compare actual vs expected — flag "indent ignored" when actual ≈ page_margin

Output: pipeline_data/bug_b_indent_pattern.json with all rows.

Run: python tools/metrics/measure_bug_b_indent_pattern.py
"""
import json
import os
import re
import sys
import zipfile

import win32com.client

WD_HPOS = 5
WD_VPOS = 6
WD_PAGE = 3

DOCS = [
    ("191cb5254cb2", "FAIL"),
    ("bd90b00ab7a7", "FAIL"),
    ("cb8be715d839", "FAIL"),
    ("1636d28e2c46", "FAIL"),
    ("de6e32b5960b", "FAIL"),
    ("b35123fe8efc", "PASS"),
    ("b5f706e9f6ad", "PASS"),
    ("d6fd9a516382", "PASS"),
]

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_DIR = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
OUT = os.path.join(REPO, "pipeline_data", "bug_b_indent_pattern.json")


def find_docx(doc_id):
    for f in os.listdir(DOCX_DIR):
        if f.startswith(doc_id) and f.endswith(".docx"):
            return os.path.join(DOCX_DIR, f)
    return None


def extract_para_indents_from_xml(docx_path):
    """Walk document.xml and yield (para_idx_1based, indent_props, leading_text_info)
    only for paragraphs with BOTH firstLineChars AND firstLine in <w:ind/>.
    """
    with zipfile.ZipFile(docx_path) as zf:
        doc_xml = zf.read("word/document.xml").decode("utf-8")

    paras = re.finditer(r"<w:p\b[^>]*>(.*?)</w:p>", doc_xml, re.DOTALL)
    out = []
    para_idx = 0
    for pm in paras:
        para_idx += 1
        body = pm.group(1)
        # Find <w:ind/>
        ind_m = re.search(r"<w:ind\b([^/]*)/>", body)
        if not ind_m:
            continue
        ind_attrs = ind_m.group(1)
        attrs = dict(re.findall(r'w:(\w+)="([^"]*)"', ind_attrs))
        if "firstLineChars" not in attrs or "firstLine" not in attrs:
            continue
        # Extract pStyle if any
        pstyle_m = re.search(r'<w:pStyle\s+w:val="([^"]+)"', body)
        pstyle = pstyle_m.group(1) if pstyle_m else None
        # Extract first run's font size
        sz_m = re.search(r'<w:sz\s+w:val="([^"]+)"', body)
        sz_hp = int(sz_m.group(1)) if sz_m else 21  # half-points (default 10.5)
        # Concatenate all <w:t> text
        ts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body)
        full_text = "".join(ts)
        leading_ws_count = 0
        for c in full_text:
            if c in (" ", "　"):
                leading_ws_count += 1
            else:
                break
        # Whether has snapToGrid=0
        snap_off = '<w:snapToGrid w:val="0"/>' in body
        out.append({
            "para_idx": para_idx,
            "ind": attrs,
            "pStyle": pstyle,
            "sz_hp": sz_hp,
            "leading_ws": leading_ws_count,
            "text_preview": full_text[:60],
            "text_len": len(full_text),
            "snap_to_grid_off": snap_off,
        })
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    all_rows = []
    try:
        for doc_id, expected in DOCS:
            docx = find_docx(doc_id)
            if not docx:
                print(f"{doc_id}: NO DOCX")
                continue
            candidates = extract_para_indents_from_xml(docx)
            if not candidates:
                print(f"{doc_id}: 0 candidate paras")
                continue
            print(f"{doc_id} ({expected}): {len(candidates)} candidate paras")
            doc = word.Documents.Open(docx, ReadOnly=True)
            try:
                # Get section margins (first section)
                sec1 = doc.Sections(1)
                pm_left_pt = sec1.PageSetup.LeftMargin
                pm_top_pt = sec1.PageSetup.TopMargin
                # Sample candidates: first 3 + any with L>0 (more interesting)
                # + bd90b00 pi=24 (the 統計センター para we already know about)
                pi_seen = set()
                sampled = []
                for cand in candidates:
                    if cand["para_idx"] == 24 and doc_id == "bd90b00ab7a7":
                        sampled.append(cand)
                        pi_seen.add(cand["para_idx"])
                # First 3 candidates
                for cand in candidates[:3]:
                    if cand["para_idx"] not in pi_seen:
                        sampled.append(cand)
                        pi_seen.add(cand["para_idx"])
                # Plus any with leading whitespace > 5
                for cand in candidates:
                    if cand["leading_ws"] >= 5 and cand["para_idx"] not in pi_seen and len(sampled) < 8:
                        sampled.append(cand)
                        pi_seen.add(cand["para_idx"])
                for cand in sampled:
                    pi = cand["para_idx"]
                    try:
                        para = doc.Paragraphs(pi)
                        rng = para.Range
                        first = doc.Range(rng.Start, rng.Start)
                        actual_x = first.Information(WD_HPOS)
                        actual_y = first.Information(WD_VPOS)
                        actual_pg = first.Information(WD_PAGE)
                        # Compute expected x
                        ind = cand["ind"]
                        left_tw = int(ind.get("left", "0"))
                        firstline_tw = int(ind.get("firstLine", "0"))
                        left_pt = left_tw / 20.0
                        firstline_pt = firstline_tw / 20.0
                        # Expected x = pm_left + left + firstLine
                        expected_x = pm_left_pt + left_pt + firstline_pt
                        # Also expected if firstLine ignored
                        expected_no_first = pm_left_pt + left_pt
                        # Diff
                        diff_full = actual_x - expected_x
                        diff_no_first = actual_x - expected_no_first
                        # Classify
                        if abs(diff_full) < 1.0:
                            classification = "FULL_INDENT_APPLIED"
                        elif abs(diff_no_first) < 1.0:
                            classification = "FIRSTLINE_IGNORED"
                        else:
                            classification = f"OTHER({diff_full:+.2f}/{diff_no_first:+.2f})"
                        row = {
                            "doc_id": doc_id,
                            "expected_status": expected,
                            "para_idx": pi,
                            "pStyle": cand["pStyle"],
                            "sz_hp": cand["sz_hp"],
                            "snap_to_grid_off": cand["snap_to_grid_off"],
                            "leading_ws": cand["leading_ws"],
                            "ind_left_tw": left_tw,
                            "ind_firstLine_tw": firstline_tw,
                            "ind_firstLineChars": ind.get("firstLineChars"),
                            "page_margin_left_pt": round(pm_left_pt, 3),
                            "actual_x_pt": round(actual_x, 3),
                            "actual_y_pt": round(actual_y, 3),
                            "actual_page": actual_pg,
                            "expected_x_full_indent": round(expected_x, 3),
                            "expected_x_no_first": round(expected_no_first, 3),
                            "diff_full": round(diff_full, 3),
                            "diff_no_first": round(diff_no_first, 3),
                            "classification": classification,
                            "text_preview": cand["text_preview"][:30],
                        }
                        all_rows.append(row)
                        print(
                            f"  pi={pi:3d} pStyle={cand['pStyle']!s:<6} "
                            f"L={left_tw:>4}tw FL={firstline_tw:>4}tw lead_ws={cand['leading_ws']:>3}  "
                            f"x_act={actual_x:.2f} x_exp_full={expected_x:.2f} x_exp_noFL={expected_no_first:.2f}  "
                            f"=> {classification}"
                        )
                    except Exception as e:
                        print(f"  pi={pi}: ERROR {e}")
            finally:
                doc.Close(SaveChanges=False)
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"rows": all_rows}, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT} — {len(all_rows)} rows")

    # Summary by classification
    from collections import Counter
    cls = Counter(r["classification"] for r in all_rows)
    print("\nClassification summary:")
    for c, n in cls.most_common():
        print(f"  {c}: {n}")


if __name__ == "__main__":
    main()
