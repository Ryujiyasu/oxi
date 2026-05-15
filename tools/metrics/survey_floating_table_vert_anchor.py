"""Cross-doc survey: which baseline docs use <w:tblpPr w:vertAnchor="text">?

Per [[session59-3a4f9f-drift-jumps-floating-table-footprint]] the 3a4f9f
cascade is caused by Oxi NOT reserving a vertAnchor="text" floating
table's vertical footprint for body paragraphs after it. Before any
layout fix, we need the cross-doc incidence map of this OOXML pattern:

  - How many of the 55 baseline docs use vertAnchor="text" tables?
  - Of those, which are Phase 1 PASS and which are FAIL?
  - What positioning attrs (tblpY, tblpX, horzAnchor) do they carry?
  - Is the pattern correlated with current FAIL status?

If most PASS docs lack the pattern, the proposed fix has a narrow blast
radius (low regression risk). If many PASS docs use the pattern, the
proposed fix risks regression and needs more careful design.

Instrumentation only — does NOT modify oxidocs-core or change any baseline.
Phase 1 53/55 mean 0.9842 must remain unchanged.

Output: pipeline_data/ra_manual_measurements/floating_table_vertanchor_survey.json
"""
from __future__ import annotations

import json
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = r"c:\Users\ryuji\oxi-main"
DOCS_DIR = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
WORD_SUMMARY = os.path.join(REPO, "pipeline_data", "pagination_word", "_summary.json")
DIFF_SUMMARY = os.path.join(REPO, "pipeline_data", "pagination_diff", "_summary.json")
OUT = os.path.join(
    REPO,
    "pipeline_data",
    "ra_manual_measurements",
    "floating_table_vertanchor_survey.json",
)

# Match <w:tblpPr ...> tag (self-closing or with content). The attribute order
# is unpredictable, so we regex the tag then parse attrs.
TBLPPR_RE = re.compile(r"<w:tblpPr\b([^/>]*)/?>")
ATTR_RE = re.compile(r'w:(\w+)="([^"]*)"')


def load_phase1_status() -> dict[str, dict]:
    with open(DIFF_SUMMARY, encoding="utf-8") as f:
        data = json.load(f)
    return {d["doc_id"]: d for d in data["docs"]}


def load_doc_filename_map() -> dict[str, str]:
    with open(WORD_SUMMARY, encoding="utf-8") as f:
        data = json.load(f)
    return {d["doc_id"]: d["filename"] for d in data["docs"]}


def parse_tblppr_attrs(attr_str: str) -> dict[str, str]:
    return dict(ATTR_RE.findall(attr_str))


def survey_docx(docx_path: str) -> dict:
    """Return per-doc summary of tblpPr occurrences."""
    result = {
        "n_tblppr_total": 0,
        "by_vertAnchor": Counter(),
        "by_horzAnchor": Counter(),
        "occurrences": [],
    }
    try:
        with zipfile.ZipFile(docx_path) as z:
            names = [
                n for n in z.namelist() if n.startswith("word/") and n.endswith(".xml")
            ]
            for name in names:
                if name == "word/document.xml" or name.startswith("word/header") or name.startswith("word/footer"):
                    with z.open(name) as fh:
                        xml = fh.read().decode("utf-8", errors="replace")
                    for m in TBLPPR_RE.finditer(xml):
                        attrs = parse_tblppr_attrs(m.group(1))
                        result["n_tblppr_total"] += 1
                        v = attrs.get("vertAnchor", "<missing>")
                        h = attrs.get("horzAnchor", "<missing>")
                        result["by_vertAnchor"][v] += 1
                        result["by_horzAnchor"][h] += 1
                        result["occurrences"].append({
                            "xml_part": name,
                            "vertAnchor": v,
                            "horzAnchor": h,
                            "tblpY": attrs.get("tblpY"),
                            "tblpX": attrs.get("tblpX"),
                            "tblpYSpec": attrs.get("tblpYSpec"),
                            "tblpXSpec": attrs.get("tblpXSpec"),
                            "leftFromText": attrs.get("leftFromText"),
                            "rightFromText": attrs.get("rightFromText"),
                            "topFromText": attrs.get("topFromText"),
                            "bottomFromText": attrs.get("bottomFromText"),
                        })
    except (FileNotFoundError, zipfile.BadZipFile) as e:
        result["error"] = str(e)
    # Convert Counters to plain dicts for JSON output
    result["by_vertAnchor"] = dict(result["by_vertAnchor"])
    result["by_horzAnchor"] = dict(result["by_horzAnchor"])
    return result


def main() -> int:
    status = load_phase1_status()
    docmap = load_doc_filename_map()

    rows = []
    pattern_pass = []
    pattern_fail = []
    no_pattern_pass = []
    no_pattern_fail = []

    # vertAnchor=="text" is the specific 3a4f9f mechanism. Track it separately.
    text_anchor_pass = []
    text_anchor_fail = []

    for doc_id, fname in sorted(docmap.items()):
        docx_path = os.path.join(DOCS_DIR, fname)
        if not os.path.exists(docx_path):
            print(f"[skip] {doc_id}: docx not found at {docx_path}")
            continue
        survey = survey_docx(docx_path)
        st = status.get(doc_id, {})
        phase1_pass = st.get("pass")
        score = st.get("score")
        row = {
            "doc_id": doc_id,
            "filename": fname,
            "phase1_pass": phase1_pass,
            "score": score,
            **survey,
        }
        rows.append(row)

        n_text_anchor = survey["by_vertAnchor"].get("text", 0)
        if n_text_anchor > 0:
            if phase1_pass:
                text_anchor_pass.append(doc_id)
            else:
                text_anchor_fail.append(doc_id)
        if survey["n_tblppr_total"] > 0:
            if phase1_pass:
                pattern_pass.append(doc_id)
            else:
                pattern_fail.append(doc_id)
        else:
            if phase1_pass:
                no_pattern_pass.append(doc_id)
            else:
                no_pattern_fail.append(doc_id)

    summary = {
        "n_docs": len(rows),
        "n_with_any_tblppr": sum(1 for r in rows if r["n_tblppr_total"] > 0),
        "n_with_vertAnchor_text": len(text_anchor_pass) + len(text_anchor_fail),
        "text_anchor_pass_doc_ids": text_anchor_pass,
        "text_anchor_fail_doc_ids": text_anchor_fail,
        "any_pattern_pass_doc_ids": pattern_pass,
        "any_pattern_fail_doc_ids": pattern_fail,
        "no_pattern_pass_doc_ids": no_pattern_pass,
        "no_pattern_fail_doc_ids": no_pattern_fail,
        "rows": rows,
    }

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print(f"\n== Survey: vertAnchor='text' floating tables across {len(rows)} baseline docs ==")
    print(f"  docs with ANY <w:tblpPr>: {summary['n_with_any_tblppr']}")
    print(f"  docs with vertAnchor='text': {summary['n_with_vertAnchor_text']}")
    print(f"    PASS Phase 1: {len(text_anchor_pass)} -> {text_anchor_pass}")
    print(f"    FAIL Phase 1: {len(text_anchor_fail)} -> {text_anchor_fail}")
    print(f"  docs with ANY tblpPr (regardless of anchor):")
    print(f"    PASS: {len(pattern_pass)}, FAIL: {len(pattern_fail)}")
    print(f"  docs with NO tblpPr:")
    print(f"    PASS: {len(no_pattern_pass)}, FAIL: {len(no_pattern_fail)}")
    print(f"\nWrote: {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
