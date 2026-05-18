"""Fuzz Quirk Discovery — docx generator.

Generates N random docx files for fuzz-based Word vs Oxi divergence detection.

Each docx contains:
- HEAD anchor paragraph
- A table with 2-4 random rows (each 1-2 cells, each cell 1-3 paragraphs)
- ANCHOR paragraph after table

The fuzz_NNNN.docx files are output to a `fuzz_runs/{batch_name}/` directory
along with a meta.json describing the attribute choices per doc.
"""
from __future__ import annotations
import json
import os
import random
import sys
import zipfile
from dataclasses import asdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from fuzz_vocabulary import FuzzPara, FuzzCell, FuzzRow


ROOT = Path(__file__).parent.parent.parent
OUT_ROOT = Path(__file__).parent / "fuzz_runs"


CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

SECTPR = '''<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701"/>
<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/>
</w:sectPr>'''


def build_table(rng: random.Random) -> tuple[str, list[dict]]:
    """Build a random table. Returns (table_xml, meta_per_row)."""
    n_rows = rng.randint(2, 4)
    n_cells = rng.choice([1, 2])
    n_paras = rng.choice([1, 1, 2])  # bias toward 1-para cells

    rows = []
    meta_rows = []
    for r_idx in range(n_rows):
        row = FuzzRow.random(rng, n_cells=n_cells, n_paras_per_cell=n_paras)
        rows.append(row)
        meta_rows.append({
            "row_idx": r_idx,
            "tr_height": row.tr_height,
            "h_rule": row.h_rule,
            "cant_split": row.cant_split,
            "cells": [{
                "width": c.width,
                "grid_span": c.grid_span,
                "v_align": c.v_align,
                "mar_top": c.mar_top,
                "mar_bottom": c.mar_bottom,
                "paragraphs": [{
                    "before": p.before, "after": p.after,
                    "before_lines": p.before_lines, "after_lines": p.after_lines,
                    "line": p.line, "line_rule": p.line_rule,
                    "left_chars": p.left_chars, "right_chars": p.right_chars,
                    "hanging_chars": p.hanging_chars, "first_line_chars": p.first_line_chars,
                    "auto_space_de": p.auto_space_de, "auto_space_dn": p.auto_space_dn,
                    "word_wrap": p.word_wrap, "adjust_right_ind": p.adjust_right_ind,
                    "jc": p.jc,
                    "font_size": p.font_size, "character_spacing": p.character_spacing,
                    "kern": p.kern, "font_family_ea": p.font_family_ea,
                } for p in c.paragraphs],
            } for c in row.cells],
        })

    # Use max grid span * n_cells as total width estimate
    max_span = max(sum(c.grid_span for c in r.cells) for r in rows)
    total_w = max_span * 3000
    grid_cols = ''.join(f'<w:gridCol w:w="3000"/>' for _ in range(max_span))

    tbl = f'''<w:tbl>
<w:tblPr>
<w:tblW w:w="{total_w}" w:type="dxa"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid>{grid_cols}</w:tblGrid>
{"".join(r.to_xml() for r in rows)}
</w:tbl>'''
    return tbl, meta_rows


def build_docx(out_path: Path, rng: random.Random) -> dict:
    table_xml, meta_rows = build_table(rng)
    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>HEAD</w:t></w:r></w:p>
{table_xml}
<w:p><w:r><w:t>ANCHOR</w:t></w:r></w:p>
{SECTPR}
</w:body></w:document>'''

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc_xml)

    return {"doc_name": out_path.name, "rows": meta_rows}


def main(n: int = 50, seed: int = 42, batch_name: str = "alpha01"):
    rng = random.Random(seed)
    out_dir = OUT_ROOT / batch_name
    out_dir.mkdir(parents=True, exist_ok=True)
    meta_all = []
    for i in range(n):
        doc_name = f"fuzz_{i:04d}.docx"
        out_path = out_dir / doc_name
        try:
            meta = build_docx(out_path, rng)
            meta_all.append(meta)
        except Exception as e:
            print(f"FAIL {doc_name}: {e}")
    with open(out_dir / "meta.json", "w", encoding="utf-8") as f:
        json.dump(meta_all, f, indent=2, ensure_ascii=False)
    print(f"Built {len(meta_all)} fuzz docs in {out_dir}")


if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 50
    seed = int(sys.argv[2]) if len(sys.argv) > 2 else 42
    batch = sys.argv[3] if len(sys.argv) > 3 else "alpha01"
    main(n=n, seed=seed, batch_name=batch)
