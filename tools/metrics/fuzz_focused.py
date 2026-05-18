"""Focused fuzz: vary 1-3 attrs while pinning everything else.

Unlike fuzz_generate.py (everything random), focused fuzz builds a
controlled experiment. Pin a baseline configuration, then sweep a
small set of attribute values.

Example sub-batches:
- font: vary font_family_ea × font_size, fix everything else
- line: vary line × line_rule, fix font, spacing, etc.
- spacing: vary before/after/Lines combinations
"""
from __future__ import annotations
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from itertools import product
from pathlib import Path

import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
RENDERER = ROOT / "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"
OUT_ROOT = Path(__file__).parent / "fuzz_runs"


# === Baseline configuration (PINNED) ===
BASELINE_PARA = {
    "before": None, "after": None,
    "before_lines": None, "after_lines": None,
    "line": None, "line_rule": None,
    "left_chars": None, "right_chars": None,
    "hanging_chars": None, "first_line_chars": None,
    "auto_space_de": None, "auto_space_dn": None,
    "word_wrap": None, "adjust_right_ind": None,
    "jc": None,
    "font_size": None, "character_spacing": None,
    "kern": None, "font_family_ea": None,
    "text": "サンプルテキスト１２３",  # contains both CJK and ASCII digits
}

BASELINE_CELL = {
    "width": 6000, "grid_span": 1,
    "v_align": None, "mar_top": None, "mar_bottom": None,
}

BASELINE_ROW = {"tr_height": None, "h_rule": None, "cant_split": None}


def build_para_xml(p: dict) -> str:
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
    rpr = []
    if p.get("font_family_ea"): rpr.append(f'<w:rFonts w:ascii="Century" w:eastAsia="{p["font_family_ea"]}" w:hAnsi="Century"/>')
    if p.get("font_size") is not None: rpr.append(f'<w:sz w:val="{p["font_size"]}"/><w:szCs w:val="{p["font_size"]}"/>')
    if p.get("character_spacing") is not None: rpr.append(f'<w:spacing w:val="{p["character_spacing"]}"/>')
    if p.get("kern") is not None: rpr.append(f'<w:kern w:val="{p["kern"]}"/>')
    rPr = f'<w:rPr>{"".join(rpr)}</w:rPr>' if rpr else ""
    text = p.get("text", "テキスト")
    return f'<w:p>{pPr}<w:r>{rPr}<w:t>{text}</w:t></w:r></w:p>'


def build_cell_xml(c: dict) -> str:
    span = f'<w:gridSpan w:val="{c["grid_span"]}"/>' if c.get("grid_span", 1) > 1 else ""
    va = f'<w:vAlign w:val="{c["v_align"]}"/>' if c.get("v_align") else ""
    mar = []
    if c.get("mar_top") is not None: mar.append(f'<w:top w:w="{c["mar_top"]}" w:type="dxa"/>')
    if c.get("mar_bottom") is not None: mar.append(f'<w:bottom w:w="{c["mar_bottom"]}" w:type="dxa"/>')
    mar_xml = f'<w:tcMar>{"".join(mar)}</w:tcMar>' if mar else ""
    paras = "".join(build_para_xml(p) for p in c["paragraphs"])
    return f'<w:tc><w:tcPr><w:tcW w:w="{c["width"]}" w:type="dxa"/>{span}{va}{mar_xml}</w:tcPr>{paras}</w:tc>'


def build_row_xml(r: dict) -> str:
    trh_parts = []
    if r.get("tr_height") is not None:
        rule = f' w:hRule="{r["h_rule"]}"' if r.get("h_rule") else ""
        trh_parts.append(f'<w:trHeight w:val="{r["tr_height"]}"{rule}/>')
    if r.get("cant_split"):
        trh_parts.append(f'<w:cantSplit/>')
    trPr = f'<w:trPr>{"".join(trh_parts)}</w:trPr>' if trh_parts else ""
    cells = "".join(build_cell_xml(c) for c in r["cells"])
    return f'<w:tr>{trPr}{cells}</w:tr>'


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


def build_docx(rows: list, out: Path):
    rows_xml = "".join(build_row_xml(r) for r in rows)
    max_span = max(sum(c.get("grid_span", 1) for c in r["cells"]) for r in rows)
    grid_cols = ''.join(f'<w:gridCol w:w="3000"/>' for _ in range(max_span))
    total_w = max_span * 3000
    tbl = f'''<w:tbl>
<w:tblPr><w:tblW w:w="{total_w}" w:type="dxa"/>
<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>
</w:tblBorders></w:tblPr>
<w:tblGrid>{grid_cols}</w:tblGrid>
{rows_xml}
</w:tbl>'''
    body = f'''<w:p><w:r><w:t>HEAD</w:t></w:r></w:p>
{tbl}
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


def make_one(para_overrides: dict, cell_overrides: dict = None, row_overrides: dict = None,
             n_rows: int = 2):
    """Build single-cell-per-row, n-row docx with overrides applied to ALL paras/cells/rows."""
    para = {**BASELINE_PARA, **para_overrides}
    cell = {**BASELINE_CELL, **(cell_overrides or {}), "paragraphs": [para]}
    row = {**BASELINE_ROW, **(row_overrides or {}), "cells": [cell]}
    return [row for _ in range(n_rows)]


def collapse_y(rng):
    doc = rng.Document
    return doc.Range(rng.Start, rng.Start).Information(6)


def measure_pair(word, docx: Path) -> tuple[float, float, float]:
    """Return (head_y, cell0_y, cell1_y or NaN) — pitch = cell1 - cell0 for 2-row table."""
    doc = word.Documents.Open(str(docx.absolute()), ReadOnly=True)
    try:
        head_y = None
        for p in doc.Paragraphs:
            if p.Range.Text.strip() == "HEAD":
                head_y = collapse_y(p.Range)
                break
        cell0_y = cell1_y = None
        if doc.Tables.Count > 0:
            tbl = doc.Tables(1)
            cell0_y = collapse_y(tbl.Cell(1, 1).Range)
            if tbl.Rows.Count >= 2:
                cell1_y = collapse_y(tbl.Cell(2, 1).Range)
    finally:
        doc.Close(SaveChanges=False)
    return head_y or 0, cell0_y or 0, cell1_y or 0


def measure_oxi_pair(docx: Path) -> tuple[float, float, float]:
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, "p_")
        dump = os.path.join(tmp, "layout.json")
        subprocess.run([str(RENDERER), str(docx), prefix, "--dump-layout="+dump],
                      capture_output=True, text=True, timeout=60)
        with open(dump, encoding="utf-8") as f:
            d = json.load(f)
    head_y = cell0_y = cell1_y = None
    for el in d["pages"][0].get("elements", []):
        if el.get("type") != "text": continue
        text = el.get("text", "")
        if "HEAD" in text and head_y is None:
            head_y = el["y"]
        if el.get("cell_row_idx") == 0 and el.get("cell_col_idx") == 0 and cell0_y is None:
            cell0_y = el["y"]
        if el.get("cell_row_idx") == 1 and el.get("cell_col_idx") == 0 and cell1_y is None:
            cell1_y = el["y"]
    return head_y or 0, cell0_y or 0, cell1_y or 0


def run_sweep(name: str, variants: list[tuple]):
    """variants = list of (label, para_overrides, cell_overrides, row_overrides)"""
    out_dir = OUT_ROOT / name
    out_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    results = []
    try:
        for label, p_ov, c_ov, r_ov in variants:
            rows = make_one(p_ov, c_ov, r_ov, n_rows=2)
            docx = out_dir / f"{label}.docx"
            build_docx(rows, docx)
            try:
                w_h, w_c0, w_c1 = measure_pair(word, docx)
                o_h, o_c0, o_c1 = measure_oxi_pair(docx)
                w_pitch = w_c1 - w_c0
                o_pitch = o_c1 - o_c0
                w_first_dy = w_c0 - w_h
                o_first_dy = o_c0 - o_h
                results.append({
                    "label": label,
                    "p_overrides": p_ov, "c_overrides": c_ov, "r_overrides": r_ov,
                    "word": {"head": w_h, "c0": w_c0, "c1": w_c1, "pitch": w_pitch, "first_dy": w_first_dy},
                    "oxi": {"head": o_h, "c0": o_c0, "c1": o_c1, "pitch": o_pitch, "first_dy": o_first_dy},
                    "pitch_diff": w_pitch - o_pitch,
                    "first_dy_diff": w_first_dy - o_first_dy,
                })
            except Exception as e:
                results.append({"label": label, "error": str(e)})
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    out = out_dir / "results.json"
    out.write_text(json.dumps(results, indent=2, default=str, ensure_ascii=False), encoding="utf-8")

    print(f"\n=== {name} ===")
    print(f"{'label':<35} {'W pitch':<8} {'O pitch':<8} {'pitch Δ':<8} {'first dy Δ':<10}")
    for r in results:
        if "error" in r:
            print(f"{r['label']:<35} ERROR: {r['error'][:50]}")
        else:
            print(f"{r['label']:<35} {r['word']['pitch']:<8.2f} {r['oxi']['pitch']:<8.2f} "
                  f"{r['pitch_diff']:+<8.2f} {r['first_dy_diff']:+<.2f}")


if __name__ == "__main__":
    # Default: line_rule × line sweep
    variants = []
    for lr in [None, "exact", "atLeast", "auto"]:
        for line in [None, 240, 280, 360]:
            label = f"lr={lr}_line={line}"
            variants.append((label, {"line_rule": lr, "line": line}, None, None))
    run_sweep("focused_line", variants)
