"""Fuzz Auto-Bisection — shrink high-divergence fuzz docs to minimal causative attrs.

Algorithm:
1. Take a fuzz doc with max_diff = D.
2. For each removable attribute, build a variant without it.
3. Re-measure variant; if max_diff drops by ≥0.3*D (significantly), the
   attribute is contributing.
4. Greedy: keep removing non-contributing attrs until minimal set remains.
5. Output: minimal docx + list of removed (innocent) attrs and remaining
   (causative) attrs.

This separates attributes that CAUSE divergence from those that just
co-occur in random fuzz.
"""
from __future__ import annotations
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
RENDERER = ROOT / "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"


PARA_ATTRS = [
    "before", "after", "before_lines", "after_lines",
    "line", "line_rule",
    "left_chars", "right_chars", "hanging_chars", "first_line_chars",
    "auto_space_de", "auto_space_dn", "word_wrap", "adjust_right_ind",
    "jc",
    "font_size", "character_spacing", "kern", "font_family_ea",
]
CELL_ATTRS = ["v_align", "mar_top", "mar_bottom"]
ROW_ATTRS = ["tr_height", "h_rule", "cant_split"]


def meta_to_docx_xml(meta_rows: list[dict]) -> str:
    """Reconstruct table XML from meta. Used after attribute removal."""
    rows_xml = ""
    max_span = max(sum(c["grid_span"] for c in r["cells"]) for r in meta_rows)
    for row in meta_rows:
        trh_parts = []
        if row.get("tr_height") is not None:
            rule = f' w:hRule="{row["h_rule"]}"' if row.get("h_rule") else ""
            trh_parts.append(f'<w:trHeight w:val="{row["tr_height"]}"{rule}/>')
        if row.get("cant_split"):
            trh_parts.append(f'<w:cantSplit/>')
        trPr = f'<w:trPr>{"".join(trh_parts)}</w:trPr>' if trh_parts else ""
        cells_xml = ""
        for cell in row["cells"]:
            span = f'<w:gridSpan w:val="{cell["grid_span"]}"/>' if cell["grid_span"] > 1 else ""
            va = f'<w:vAlign w:val="{cell["v_align"]}"/>' if cell.get("v_align") else ""
            mar = []
            if cell.get("mar_top") is not None:
                mar.append(f'<w:top w:w="{cell["mar_top"]}" w:type="dxa"/>')
            if cell.get("mar_bottom") is not None:
                mar.append(f'<w:bottom w:w="{cell["mar_bottom"]}" w:type="dxa"/>')
            mar_xml = f'<w:tcMar>{"".join(mar)}</w:tcMar>' if mar else ""
            paras_xml = ""
            for p in cell["paragraphs"]:
                sp_parts = []
                if p.get("before") is not None: sp_parts.append(f'w:before="{p["before"]}"')
                if p.get("after") is not None: sp_parts.append(f'w:after="{p["after"]}"')
                if p.get("before_lines") is not None: sp_parts.append(f'w:beforeLines="{p["before_lines"]}"')
                if p.get("after_lines") is not None: sp_parts.append(f'w:afterLines="{p["after_lines"]}"')
                if p.get("line") is not None: sp_parts.append(f'w:line="{p["line"]}"')
                if p.get("line_rule") is not None: sp_parts.append(f'w:lineRule="{p["line_rule"]}"')
                sp = f'<w:spacing {" ".join(sp_parts)}/>' if sp_parts else ""
                ind_parts = []
                if p.get("left_chars") is not None: ind_parts.append(f'w:leftChars="{p["left_chars"]}"')
                if p.get("right_chars") is not None: ind_parts.append(f'w:rightChars="{p["right_chars"]}"')
                if p.get("hanging_chars") is not None: ind_parts.append(f'w:hangingChars="{p["hanging_chars"]}"')
                if p.get("first_line_chars") is not None: ind_parts.append(f'w:firstLineChars="{p["first_line_chars"]}"')
                ind = f'<w:ind {" ".join(ind_parts)}/>' if ind_parts else ""
                flags = []
                if p.get("auto_space_de") is not None: flags.append(f'<w:autoSpaceDE w:val="{p["auto_space_de"]}"/>')
                if p.get("auto_space_dn") is not None: flags.append(f'<w:autoSpaceDN w:val="{p["auto_space_dn"]}"/>')
                if p.get("word_wrap") is not None: flags.append(f'<w:wordWrap w:val="{p["word_wrap"]}"/>')
                if p.get("adjust_right_ind") is not None: flags.append(f'<w:adjustRightInd w:val="{p["adjust_right_ind"]}"/>')
                jc = f'<w:jc w:val="{p["jc"]}"/>' if p.get("jc") else ""
                pPr = f'<w:pPr>{sp}{ind}{"".join(flags)}{jc}</w:pPr>'
                rpr_parts = []
                if p.get("font_family_ea") is not None: rpr_parts.append(f'<w:rFonts w:ascii="Century" w:eastAsia="{p["font_family_ea"]}" w:hAnsi="Century"/>')
                if p.get("font_size") is not None: rpr_parts.append(f'<w:sz w:val="{p["font_size"]}"/><w:szCs w:val="{p["font_size"]}"/>')
                if p.get("character_spacing") is not None: rpr_parts.append(f'<w:spacing w:val="{p["character_spacing"]}"/>')
                if p.get("kern") is not None: rpr_parts.append(f'<w:kern w:val="{p["kern"]}"/>')
                rPr = f'<w:rPr>{"".join(rpr_parts)}</w:rPr>' if rpr_parts else ""
                text = p.get("text", "サンプルテキスト")
                paras_xml += f'<w:p>{pPr}<w:r>{rPr}<w:t>{text}</w:t></w:r></w:p>'
            cells_xml += (f'<w:tc><w:tcPr><w:tcW w:w="{cell["width"]}" w:type="dxa"/>'
                         f'{span}{va}{mar_xml}</w:tcPr>{paras_xml}</w:tc>')
        rows_xml += f'<w:tr>{trPr}{cells_xml}</w:tr>'

    grid_cols = ''.join(f'<w:gridCol w:w="3000"/>' for _ in range(max_span))
    total_w = max_span * 3000
    return f'''<w:tbl>
<w:tblPr>
<w:tblW w:w="{total_w}" w:type="dxa"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid>{grid_cols}</w:tblGrid>
{rows_xml}
</w:tbl>'''


CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>'''
SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat></w:settings>'''


def build_docx(meta_rows: list, out: Path):
    table_xml = meta_to_docx_xml(meta_rows)
    body = f'''<w:p><w:r><w:t>HEAD</w:t></w:r></w:p>
{table_xml}
<w:p><w:r><w:t>ANCHOR</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701"/>
<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/></w:sectPr>'''
    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>{body}</w:body></w:document>'''
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc_xml)


def collapse_y(rng):
    doc = rng.Document
    return doc.Range(rng.Start, rng.Start).Information(6)


def measure_doc(word, docx: Path) -> float:
    """Return max |Word - Oxi| delta_y across all cells."""
    # Word
    doc = word.Documents.Open(str(docx.absolute()), ReadOnly=True)
    try:
        head_y = anchor_y = None
        for p in doc.Paragraphs:
            t = p.Range.Text.strip()
            if t == "HEAD":
                head_y = collapse_y(p.Range)
            elif t == "ANCHOR":
                anchor_y = collapse_y(p.Range)
        word_cells = []
        if doc.Tables.Count > 0:
            tbl = doc.Tables(1)
            for r in range(1, tbl.Rows.Count + 1):
                row = tbl.Rows(r)
                for c in range(1, row.Cells.Count + 1):
                    cell = row.Cells(c)
                    y = collapse_y(cell.Range)
                    word_cells.append(((r-1, c-1), y))
    finally:
        doc.Close(SaveChanges=False)

    # Oxi
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, "p_")
        dump = os.path.join(tmp, "layout.json")
        proc = subprocess.run([str(RENDERER), str(docx), prefix, "--dump-layout="+dump],
                              capture_output=True, text=True, timeout=60)
        if proc.returncode != 0:
            return float("nan")
        with open(dump, encoding="utf-8") as f:
            d = json.load(f)
    o_head = None
    oxi_cells = {}
    for el in d["pages"][0].get("elements", []):
        if el.get("type") != "text":
            continue
        text = el.get("text", "")
        if "HEAD" in text and o_head is None:
            o_head = el["y"]
        elif el.get("cell_row_idx") is not None:
            key = (el["cell_row_idx"], el["cell_col_idx"])
            if key not in oxi_cells or el["y"] < oxi_cells[key]:
                oxi_cells[key] = el["y"]

    if head_y is None or o_head is None:
        return float("nan")
    max_diff = 0
    for key, w_y in word_cells:
        o_y = oxi_cells.get(key)
        if o_y is None:
            continue
        w_dy = w_y - head_y
        o_dy = o_y - o_head
        d = abs(w_dy - o_dy)
        if d > max_diff:
            max_diff = d
    return max_diff


def remove_attr(meta_rows: list, attr_path: tuple) -> list:
    """Remove a single attribute from a copy of meta_rows.
    attr_path: ("para", row_idx, cell_idx, para_idx, attr_name) or
               ("cell", row_idx, cell_idx, attr_name) or
               ("row", row_idx, attr_name)
    """
    copy = deepcopy(meta_rows)
    if attr_path[0] == "para":
        _, ri, ci, pi, name = attr_path
        copy[ri]["cells"][ci]["paragraphs"][pi][name] = None
    elif attr_path[0] == "cell":
        _, ri, ci, name = attr_path
        copy[ri]["cells"][ci][name] = None
    elif attr_path[0] == "row":
        _, ri, name = attr_path
        copy[ri][name] = None
    return copy


def list_attr_paths(meta_rows: list) -> list[tuple]:
    """List all (path, name) tuples for non-None attributes."""
    paths = []
    for ri, row in enumerate(meta_rows):
        for name in ROW_ATTRS:
            if row.get(name) is not None:
                paths.append(("row", ri, name))
        for ci, cell in enumerate(row["cells"]):
            for name in CELL_ATTRS:
                if cell.get(name) is not None:
                    paths.append(("cell", ri, ci, name))
            for pi, p in enumerate(cell["paragraphs"]):
                for name in PARA_ATTRS:
                    if p.get(name) is not None:
                        paths.append(("para", ri, ci, pi, name))
    return paths


def bisect(meta_rows: list, work_dir: Path, baseline_diff: float, min_keep_frac: float = 0.5):
    """Greedy bisection: remove attrs one by one if removal doesn't drop diff much.

    Returns (minimal_meta, removed_attrs, kept_attrs).
    """
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    current = deepcopy(meta_rows)
    removed = []
    kept = []
    try:
        paths = list_attr_paths(current)
        print(f"  Initial {len(paths)} attrs, baseline diff={baseline_diff:.2f}pt")
        for i, ap in enumerate(paths):
            # Re-list paths since current may have changed (paths still valid by index)
            trial = remove_attr(current, ap)
            trial_path = work_dir / f"trial_{i:03d}.docx"
            build_docx(trial, trial_path)
            try:
                d = measure_doc(word, trial_path)
            except Exception as e:
                d = float("nan")
            os.remove(trial_path)
            # If diff drops significantly (≥30%), this attr was contributing — KEEP
            # If diff stays ≥ min_keep_frac * baseline, this attr is innocent — REMOVE
            if d != d:  # NaN
                kept.append(ap)
                continue
            if d >= baseline_diff * min_keep_frac:
                # Removing didn't hurt much — innocent
                current = trial
                removed.append((ap, d))
            else:
                kept.append((ap, d))
        print(f"  Final: {len(kept)} kept (causative), {len(removed)} removed (innocent)")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
    return current, removed, kept


def main(batch: str, top_n: int = 3):
    batch_dir = Path(__file__).parent / "fuzz_runs" / batch
    div = json.loads((batch_dir / "divergences.json").read_text(encoding="utf-8"))
    meta = {m["doc_name"]: m for m in json.loads((batch_dir / "meta.json").read_text(encoding="utf-8"))}

    # Filter clean docs (single-page consistent)
    clean = [d for d in div
             if d.get("anchor_diff") is not None
             and abs(d.get("anchor_diff", 0)) < 50
             and d.get("max_diff", 0) > 30]
    clean_sorted = sorted(clean, key=lambda x: -x.get("max_diff", 0))[:top_n]

    out_dir = batch_dir / "bisect"
    out_dir.mkdir(exist_ok=True)

    for d in clean_sorted:
        doc_name = d["doc"]
        baseline_diff = d["max_diff"]
        print(f"\n=== Bisecting {doc_name} (baseline_diff={baseline_diff:.2f}pt) ===")
        m = meta[doc_name]
        minimal_meta, removed, kept = bisect(m["rows"], out_dir, baseline_diff)
        # Save minimal
        out = out_dir / doc_name.replace(".docx", "_minimal.docx")
        build_docx(minimal_meta, out)
        # Save report
        report = {
            "doc": doc_name,
            "baseline_diff": baseline_diff,
            "minimal_path": str(out),
            "kept_attrs (causative)": [
                {"path": list(k[0]) if isinstance(k, tuple) else list(k), "diff_after_removing": k[1] if isinstance(k, tuple) and len(k) > 1 else None}
                for k in kept
            ],
            "removed_attrs (innocent)": [
                {"path": list(r[0]), "diff_after_removing": r[1]}
                for r in removed
            ],
        }
        (out_dir / doc_name.replace(".docx", "_report.json")).write_text(
            json.dumps(report, indent=2, default=str), encoding="utf-8"
        )
        print(f"  Kept (causative):")
        for k in kept[:15]:
            if isinstance(k, tuple):
                print(f"    {k[0]} — diff={k[1] if len(k) > 1 else 'n/a'}")
            else:
                print(f"    {k}")


if __name__ == "__main__":
    batch = sys.argv[1] if len(sys.argv) > 1 else "alpha01"
    top_n = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    main(batch=batch, top_n=top_n)
