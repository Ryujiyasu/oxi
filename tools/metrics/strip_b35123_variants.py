"""Strip b35123 element-by-element to localize the cell-charGrid bug trigger.

S109e: 15 minimal repros could not reproduce the bug. S109f (this script):
take actual b35123 and strip ONE element at a time. Whichever strip makes
the bug disappear identifies the trigger.

Strip variants (each independent):
- S0: baseline (no strip)
- S1: strip rows 3-13 (keep only row 1 + row 2)
- S2: strip rows 3-13 AND row 2 (keep only row 1)
- S3: strip preceding paragraphs (table becomes first body element)
- S4: strip succeeding paragraphs after the first table
- S5: strip second table entirely
- S6: simplify tblLook (use minimal "0000")
- S7: strip table style "af" reference (use no style)
- S8: strip settings.xml
- S9: strip theme1.xml
- S10: strip numbering.xml
- S11: ALL of S1-S10 combined
"""
import os
import re
import zipfile
from io import BytesIO

SRC = os.path.abspath("tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx")
OUT_DIR = os.path.abspath("tools/metrics/b35123_strip_variants")
os.makedirs(OUT_DIR, exist_ok=True)


def read_docx(path):
    """Read all parts of docx as dict {name: bytes}."""
    parts = {}
    with zipfile.ZipFile(path, "r") as z:
        for name in z.namelist():
            parts[name] = z.read(name)
    return parts


def write_docx(parts, out_path):
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def strip_table_rows(doc_xml: str, table_idx: int, keep_first_n_rows: int) -> str:
    """Keep only first N rows of the table_idx-th table."""
    tables = list(re.finditer(r"<w:tbl\b[^>]*>.*?</w:tbl>", doc_xml, re.DOTALL))
    if table_idx >= len(tables):
        return doc_xml
    tbl_m = tables[table_idx]
    tbl_xml = tbl_m.group(0)
    # Find rows
    rows = list(re.finditer(r"<w:tr\b[^>]*>.*?</w:tr>", tbl_xml, re.DOTALL))
    if len(rows) <= keep_first_n_rows:
        return doc_xml
    # Build new table with only first N rows
    new_tbl = tbl_xml[: rows[0].start()] + "".join(r.group(0) for r in rows[:keep_first_n_rows]) + tbl_xml[rows[-1].end() :]
    return doc_xml[: tbl_m.start()] + new_tbl + doc_xml[tbl_m.end() :]


def strip_table(doc_xml: str, table_idx: int) -> str:
    tables = list(re.finditer(r"<w:tbl\b[^>]*>.*?</w:tbl>", doc_xml, re.DOTALL))
    if table_idx >= len(tables):
        return doc_xml
    tbl_m = tables[table_idx]
    return doc_xml[: tbl_m.start()] + doc_xml[tbl_m.end() :]


def strip_paragraphs_before_first_table(doc_xml: str) -> str:
    """Replace everything between <w:body> and the first <w:tbl> with empty."""
    body_start = doc_xml.find("<w:body>")
    tbl_start = doc_xml.find("<w:tbl")
    if body_start < 0 or tbl_start < 0 or tbl_start < body_start:
        return doc_xml
    return doc_xml[: body_start + len("<w:body>")] + doc_xml[tbl_start:]


def strip_paragraphs_after_first_table(doc_xml: str) -> str:
    """After the first table, drop all paragraphs until <w:sectPr> or </w:body>."""
    tables = list(re.finditer(r"<w:tbl\b[^>]*>.*?</w:tbl>", doc_xml, re.DOTALL))
    if not tables:
        return doc_xml
    tbl_end = tables[0].end()
    # Find sectPr or end of body
    sect_m = re.search(r"<w:sectPr\b[^>]*>", doc_xml[tbl_end:])
    body_close = doc_xml.find("</w:body>", tbl_end)
    if sect_m:
        cut = tbl_end + sect_m.start()
    elif body_close >= 0:
        cut = body_close
    else:
        return doc_xml
    return doc_xml[:tbl_end] + doc_xml[cut:]


def modify_tbllook(doc_xml: str, table_idx: int, new_val: str) -> str:
    tables = list(re.finditer(r"<w:tbl\b[^>]*>.*?</w:tbl>", doc_xml, re.DOTALL))
    if table_idx >= len(tables):
        return doc_xml
    tbl_m = tables[table_idx]
    tbl_xml = tbl_m.group(0)
    new_tbl = re.sub(r'<w:tblLook[^/]*/>', f'<w:tblLook w:val="{new_val}"/>', tbl_xml, count=1)
    return doc_xml[: tbl_m.start()] + new_tbl + doc_xml[tbl_m.end() :]


def strip_table_style(doc_xml: str, table_idx: int) -> str:
    tables = list(re.finditer(r"<w:tbl\b[^>]*>.*?</w:tbl>", doc_xml, re.DOTALL))
    if table_idx >= len(tables):
        return doc_xml
    tbl_m = tables[table_idx]
    tbl_xml = tbl_m.group(0)
    new_tbl = re.sub(r'<w:tblStyle[^/]*/>', '', tbl_xml, count=1)
    return doc_xml[: tbl_m.start()] + new_tbl + doc_xml[tbl_m.end() :]


def apply_variant(parts: dict, variant: str) -> dict:
    parts = {k: v for k, v in parts.items()}
    doc_xml = parts["word/document.xml"].decode("utf-8")
    if variant == "S0":
        pass
    elif variant == "S1":
        doc_xml = strip_table_rows(doc_xml, 0, keep_first_n_rows=2)
    elif variant == "S2":
        doc_xml = strip_table_rows(doc_xml, 0, keep_first_n_rows=1)
    elif variant == "S3":
        doc_xml = strip_paragraphs_before_first_table(doc_xml)
    elif variant == "S4":
        doc_xml = strip_paragraphs_after_first_table(doc_xml)
    elif variant == "S5":
        # strip the SECOND table; index 1
        doc_xml = strip_table(doc_xml, 1)
    elif variant == "S6":
        doc_xml = modify_tbllook(doc_xml, 0, "0000")
    elif variant == "S7":
        doc_xml = strip_table_style(doc_xml, 0)
    elif variant == "S8":
        if "word/settings.xml" in parts:
            del parts["word/settings.xml"]
            # Also remove the reference from document.xml.rels and [Content_Types].xml
            ct = parts["[Content_Types].xml"].decode("utf-8")
            ct = re.sub(r'<Override PartName="/word/settings\.xml"[^/]*/>', "", ct)
            parts["[Content_Types].xml"] = ct.encode("utf-8")
            if "word/_rels/document.xml.rels" in parts:
                rels = parts["word/_rels/document.xml.rels"].decode("utf-8")
                rels = re.sub(r'<Relationship[^>]+Target="settings\.xml"[^/]*/>', "", rels)
                parts["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    elif variant == "S9":
        if "word/theme/theme1.xml" in parts:
            del parts["word/theme/theme1.xml"]
            ct = parts["[Content_Types].xml"].decode("utf-8")
            ct = re.sub(r'<Override PartName="/word/theme/theme1\.xml"[^/]*/>', "", ct)
            parts["[Content_Types].xml"] = ct.encode("utf-8")
            if "word/_rels/document.xml.rels" in parts:
                rels = parts["word/_rels/document.xml.rels"].decode("utf-8")
                rels = re.sub(r'<Relationship[^>]+Target="theme/theme1\.xml"[^/]*/>', "", rels)
                parts["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    elif variant == "S10":
        if "word/numbering.xml" in parts:
            del parts["word/numbering.xml"]
            ct = parts["[Content_Types].xml"].decode("utf-8")
            ct = re.sub(r'<Override PartName="/word/numbering\.xml"[^/]*/>', "", ct)
            parts["[Content_Types].xml"] = ct.encode("utf-8")
            if "word/_rels/document.xml.rels" in parts:
                rels = parts["word/_rels/document.xml.rels"].decode("utf-8")
                rels = re.sub(r'<Relationship[^>]+Target="numbering\.xml"[^/]*/>', "", rels)
                parts["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    elif variant == "S11":
        # apply all of S1-S10
        for v in ["S1", "S3", "S4", "S5", "S6", "S7", "S8", "S9", "S10"]:
            parts = apply_variant(parts, v)
            doc_xml = parts["word/document.xml"].decode("utf-8")
    parts["word/document.xml"] = doc_xml.encode("utf-8")
    return parts


def main():
    parts = read_docx(SRC)
    variants = ["S0", "S1", "S2", "S3", "S4", "S5", "S6", "S7", "S8", "S9", "S10", "S11"]
    for v in variants:
        try:
            new_parts = apply_variant(parts, v)
            out_path = os.path.join(OUT_DIR, f"b35_{v}.docx")
            write_docx(new_parts, out_path)
            print(f"Built {out_path}")
        except Exception as e:
            print(f"FAIL {v}: {e}")


if __name__ == "__main__":
    main()
