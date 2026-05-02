"""§13.6 Class A — measure shifted-paragraph-in-table share for 5 docs.

For each of {04b88e, 6514f, d4d126, 34140, 459f}:
  1. Parse postR41 pagination diff for shifted paragraphs (page_delta != 0)
  2. Extract paragraph positions from document.xml
  3. Determine which shifted paragraphs are inside <w:tbl>
  4. Compute in-table share

Per-doc structural features (n_tables, lineRule mix, text_scale,
drawings, textboxes) for context.
"""
import json
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = "C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
DIFFS = "C:/Users/ryuji/oxi-main/pipeline_data/pagination_diff_postR41"

DOCS = ["04b88e7e0b25", "6514f214e482", "d4d126dfe1d9",
        "34140b9c5662", "459f05f1e877"]


def find_docx(stem):
    for f in os.listdir(DOCX):
        if f.startswith(stem) and f.endswith(".docx"):
            return os.path.join(DOCX, f)
    return None


def analyze_doc(stem):
    docx_path = find_docx(stem)
    diff_path = os.path.join(DIFFS, f"{stem}.json")
    if not docx_path or not os.path.exists(diff_path):
        return None

    with open(diff_path, encoding="utf-8") as f:
        diff = json.load(f)
    shifted = [m['word_i'] for m in diff['matches']
               if m.get('page_delta') and m['page_delta'] != 0]

    with zipfile.ZipFile(docx_path) as zf:
        doc_xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
        try:
            settings_xml = zf.read("word/settings.xml").decode("utf-8", errors="replace")
        except KeyError:
            settings_xml = ""

    body = re.search(r"<w:body>(.*)</w:body>", doc_xml, re.DOTALL).group(1)

    # Table ranges
    tbl_ranges = [(m.start(), m.end())
                  for m in re.finditer(r"<w:tbl[\s>].*?</w:tbl>", body, re.DOTALL)]

    # Paragraph positions
    para_positions = [(m.start(), m.end(), idx + 1)
                      for idx, m in enumerate(re.finditer(r"<w:p\b[^>]*>.*?</w:p>",
                                                          body, re.DOTALL))]

    in_tbl_count = 0
    not_in_tbl = []
    not_found = []
    for sp in shifted:
        found = False
        for pos_s, pos_e, pi in para_positions:
            if pi == sp:
                found = True
                in_tbl = any(tr[0] <= pos_s < tr[1] for tr in tbl_ranges)
                if in_tbl:
                    in_tbl_count += 1
                else:
                    not_in_tbl.append(sp)
                break
        if not found:
            not_found.append(sp)

    # Structural features
    n_paras = len(para_positions)
    n_tbls = len(tbl_ranges)
    n_drawings = len(re.findall(r"<w:drawing\b", doc_xml))
    n_textboxes = len(re.findall(r"<w:txbxContent\b", doc_xml))
    n_floats = len(re.findall(r"<w:tblpPr", doc_xml))
    n_w_w = len(re.findall(r'<w:w w:val="\d+"', doc_xml))
    n_exact = len(re.findall(r'w:lineRule="exact"', doc_xml))
    n_atLeast = len(re.findall(r'w:lineRule="atLeast"', doc_xml))
    n_auto = len(re.findall(r'w:lineRule="auto"', doc_xml))
    cSC = ('compressPunctuation' in settings_xml and 'compressPunctuation') or \
          ('doNotCompress' in settings_xml and 'doNotCompress') or 'unspec'

    return {
        "stem": stem,
        "n_paras": n_paras,
        "n_tables": n_tbls,
        "n_drawings": n_drawings,
        "n_textboxes": n_textboxes,
        "n_float_tables": n_floats,
        "cSC": cSC,
        "lineRule": {"exact": n_exact, "atLeast": n_atLeast, "auto": n_auto},
        "n_w_w": n_w_w,
        "score": diff.get("score"),
        "n_shifted": len(shifted),
        "n_shifted_in_table": in_tbl_count,
        "n_shifted_not_in_table": len(not_in_tbl),
        "n_shifted_not_found": len(not_found),
        "in_table_share": (in_tbl_count / len(shifted)) if shifted else None,
    }


def main():
    print(f"{'doc':>14} {'rank':>5} {'shifted':>7} {'inTbl':>6}"
          f" {'share':>7}  {'tbls':>4} {'draw':>4} {'tbx':>3} {'fl':>2}"
          f"  exact/atL/auto  text_scale  cSC")

    # Approximate ranks from baseline
    ranks = {"04b88e7e0b25": 8, "6514f214e482": 12, "d4d126dfe1d9": 9,
             "34140b9c5662": 10, "459f05f1e877": 11}

    summary = {}
    for stem in DOCS:
        r = analyze_doc(stem)
        if r is None:
            print(f"  {stem}: NOT FOUND")
            continue
        summary[stem] = r
        rank = ranks.get(stem, "?")
        share = r["in_table_share"]
        share_str = f"{share*100:>5.1f}%" if share is not None else "  N/A"
        lr = r["lineRule"]
        cSC_short = "Y" if "compress" in r["cSC"] else "N"
        print(f"  {stem:>14} {rank:>5} {r['n_shifted']:>7d}"
              f" {r['n_shifted_in_table']:>6d} {share_str:>7s}"
              f"  {r['n_tables']:>4d} {r['n_drawings']:>4d}"
              f" {r['n_textboxes']:>3d} {r['n_float_tables']:>2d}"
              f"  {lr['exact']:>3d}/{lr['atLeast']:>2d}/{lr['auto']:>2d}"
              f"     {r['n_w_w']:>4d}      {cSC_short}")

    # Aggregate stats
    print()
    total_shifted = sum(r['n_shifted'] for r in summary.values())
    total_in_tbl = sum(r['n_shifted_in_table'] for r in summary.values())
    print(f"Aggregate: {total_in_tbl}/{total_shifted} "
          f"({100 * total_in_tbl / total_shifted:.1f}%) shifted paras in tables "
          f"across {len(summary)} docs")

    out_path = os.path.join(os.path.dirname(__file__), "output",
                            "class_A_table_share_5docs.json")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {out_path}")


if __name__ == "__main__":
    main()
