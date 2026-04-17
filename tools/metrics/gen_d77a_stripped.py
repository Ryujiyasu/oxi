"""Subtractive bisection: make copies of d77a with content stripped.

Strategy: start from full d77a, progressively remove paragraphs / settings
until ・=9.5pt compression stops. The last property whose removal changes
the measurement is the trigger.

Variants:
  S1: d77a with ONLY para 28 kept (all other body paras removed, keep sectPr)
  S2: S1 but also keep para 27 (preceding para) — maybe prior para state matters
  S3: d77a with all tables removed (keep all body paras)
  S4: d77a with footnotes removed
  S5: d77a with settings.xml replaced by minimal (no compat)
"""
import os, shutil, re, zipfile, tempfile

SRC = os.path.abspath(
    r"tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx"
)
OUT_DIR = os.path.abspath(r"pipeline_data")


def rewrite_docx(src, dst, transforms):
    """transforms: dict {filename: callable(data: bytes) -> bytes}"""
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item in transforms:
                    data = transforms[item](data)
                zout.writestr(item, data)


def strip_body_keep_only_para28(data):
    """Replace body content with ONLY para 28 + sectPr."""
    xml = data.decode("utf-8")
    # Extract sectPr
    sectpr_m = re.search(r'<w:sectPr.*?</w:sectPr>', xml, re.DOTALL)
    sectpr = sectpr_m.group(0) if sectpr_m else ''

    # Find para 28 — the one starting with ・利用規約名
    # Find all p blocks at top level of body
    body_m = re.search(r'(<w:body>)(.*)(</w:body>)', xml, re.DOTALL)
    if not body_m:
        return data
    body_start, body_content, body_end = body_m.group(1), body_m.group(2), body_m.group(3)

    # Find para 28 block
    pos = 0
    para28 = None
    while True:
        m = re.search(r'<w:p[ >]', body_content[pos:])
        if not m: break
        start = pos + m.start()
        # Find matching </w:p> at same depth
        depth = 1
        i = start + len(m.group(0))
        while i < len(body_content) and depth > 0:
            if body_content[i:i+5] == '<w:p>' or body_content[i:i+5] == '<w:p ':
                depth += 1; i += 5
            elif body_content[i:i+6] == '</w:p>':
                depth -= 1; i += 6
            else:
                i += 1
        end = i
        block = body_content[start:end]
        if '・利用規約名' in block:
            para28 = block
            break
        pos = end

    if para28 is None:
        print("WARN: para 28 not found in body")
        return data

    new_body = body_start + para28 + sectpr + body_end
    new_xml = xml[:body_m.start()] + new_body + xml[body_m.end():]
    return new_xml.encode("utf-8")


def strip_all_tables(data):
    """Remove all <w:tbl> blocks."""
    xml = data.decode("utf-8")
    # Simple non-nested w:tbl stripping (d77a tables don't nest)
    while True:
        new_xml = re.sub(r'<w:tbl>.*?</w:tbl>', '', xml, count=1, flags=re.DOTALL)
        if new_xml == xml: break
        xml = new_xml
    return xml.encode("utf-8")


def strip_compat(data):
    """Replace w:compat block with minimal (only useFELayout)."""
    xml = data.decode("utf-8")
    xml = re.sub(r'<w:compat>.*?</w:compat>', '<w:compat><w:useFELayout/></w:compat>', xml, flags=re.DOTALL)
    return xml.encode("utf-8")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    S1 = os.path.join(OUT_DIR, "d77a_S1_only_para28.docx")
    rewrite_docx(SRC, S1, {"word/document.xml": strip_body_keep_only_para28})
    print(f"[S1] {S1}")

    S3 = os.path.join(OUT_DIR, "d77a_S3_no_tables.docx")
    rewrite_docx(SRC, S3, {"word/document.xml": strip_all_tables})
    print(f"[S3] {S3}")

    S5 = os.path.join(OUT_DIR, "d77a_S5_minimal_compat.docx")
    rewrite_docx(SRC, S5, {"word/settings.xml": strip_compat})
    print(f"[S5] {S5}")


if __name__ == "__main__":
    main()
