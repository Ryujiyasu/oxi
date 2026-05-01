"""Augment d77a chars-indent extraction with structural context:
- Is paragraph inside a <w:tbl>? At what nesting depth?
- Is paragraph inside a textbox / drawing?
- Style chain: pStyle → basedOn chain
- COM-reported LeftIndent vs predicted twip-based / chars-based / leftChars-priority
"""
import json, re, zipfile, sys
from pathlib import Path

DOC = Path(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")
EXTRACT_JSON = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\d77a_chars_indent_extract.json")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\d77a_context_overlay.json")
CHAR_WIDTH = 10.5  # docDefault.sz=21 → 10.5pt


def parse_paragraph_contexts(docx_path):
    """For each <w:p> in document.xml, determine:
    - body | table_cell | textbox | header | footer
    - table_depth (0 if body, 1 if outer cell, 2 if nested cell)
    """
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")

    contexts = []
    # We track tbl depth as we walk
    pos = 0
    tbl_depth = 0
    in_textbox = False
    # Find all relevant tags in order: <w:p, <w:tbl, </w:tbl>, <v:textbox or <w:txbxContent
    tag_re = re.compile(r'<w:p\b|</w:tbl>|<w:tbl\b|<w:txbxContent\b|</w:txbxContent>')
    for m in tag_re.finditer(xml):
        tag = m.group(0)
        if tag == '<w:p':
            # Determine context based on current state
            ctx = "body"
            if in_textbox: ctx = "textbox"
            elif tbl_depth > 0: ctx = f"table_d{tbl_depth}"
            contexts.append({"context": ctx, "tbl_depth": tbl_depth, "in_textbox": in_textbox})
        elif tag == '<w:tbl':
            tbl_depth += 1
        elif tag == '</w:tbl>':
            tbl_depth -= 1
        elif tag == '<w:txbxContent':
            in_textbox = True
        elif tag == '</w:txbxContent>':
            in_textbox = False
    return contexts


def main():
    contexts = parse_paragraph_contexts(DOC)
    print(f"Parsed {len(contexts)} paragraph contexts", file=sys.stderr)

    with open(EXTRACT_JSON, "r", encoding="utf-8") as f:
        prev = json.load(f)

    # Re-emit with context overlay
    out = {
        "docDefault_sz": prev.get("docDefault_sz"),
        "char_width_pt": CHAR_WIDTH,
        "pages": {},
    }

    for pg, rows in prev.get("pages", {}).items():
        for r in rows:
            idx = r["para_idx_word"]
            # XML index = word's reported index, but Word may count differently.
            # First try direct index.
            if idx - 1 < len(contexts):
                ctx = contexts[idx - 1]
            else:
                ctx = {"context": "?", "tbl_depth": 0, "in_textbox": False}
            r2 = dict(r)
            r2["xml_context"] = ctx["context"]
            r2["xml_tbl_depth"] = ctx["tbl_depth"]
            # Compute predicted indent values per §15.1.1
            ind = r.get("xml_ind") or {}
            l_tw = ind.get("left")
            l_ch = ind.get("leftChars")
            r2["pred_left_twip_pt"] = (l_tw / 20.0) if l_tw is not None else None
            r2["pred_left_chars_pt"] = (l_ch * CHAR_WIDTH / 100.0) if l_ch is not None else None
            r2["pred_left_priority_chars_pt"] = (
                (l_ch * CHAR_WIDTH / 100.0)
                if l_ch is not None else
                ((l_tw / 20.0) if l_tw is not None else None)
            )
            out["pages"].setdefault(pg, []).append(r2)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)

    # Summary by context per page
    print()
    print("=== context distribution per page ===")
    for pg in sorted(out["pages"], key=int):
        rows = out["pages"][pg]
        ctx_count = {}
        for r in rows:
            c = r["xml_context"]
            ctx_count[c] = ctx_count.get(c, 0) + 1
        print(f"p{pg}: {len(rows)} paras  contexts={ctx_count}")

    print()
    print("=== Per-page details with twip/chars priority resolution ===")
    for pg in sorted(out["pages"], key=int):
        rows = out["pages"][pg]
        print(f"\n--- p{pg} ---")
        for r in rows:
            ind = r.get("xml_ind") or {}
            ctx = r["xml_context"]
            ptw = r.get("pred_left_twip_pt")
            pch = r.get("pred_left_chars_pt")
            ppr = r.get("pred_left_priority_chars_pt")
            li = r.get("format_LI_pt")
            mm = ""
            if ptw is not None and pch is not None and abs(ptw - pch) > 0.5:
                mm = f"  twipVS_chars: {ptw:.1f}/{pch:.1f}"
            print(f"  p{r['para_idx_word']:3d} {ctx:>10s} y={r['y']:6.1f}"
                  f" LI_word={li:5.1f}"
                  f" pred_chars_pri={ppr:>5}"
                  f" {mm}"
                  f" ind_keys={sorted(ind.keys())}")


if __name__ == "__main__":
    main()
