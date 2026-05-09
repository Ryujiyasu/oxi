"""Extract 0e7af1ae8f21 paragraph 207 (xml idx) as minimal repro.

Day 31 part 29 identified paragraph 207 as the sole class B bug
violator: 9pt font + hanging indent + 103 chars.

Hypothesis: Oxi wraps to 1 line, Word wraps to 2+ lines. With SOFT_MARGIN
+0.5pt, Oxi line 1 fits on page 6 → orphan check doesn't fire (lines.len() < 2)
→ paragraph 207 incorrectly counted as on page 6 instead of Word's page 7.

Output: tools/golden-test/repros/0e7af_class_b/CB_V200_para207_only.docx
"""
from __future__ import annotations
import os, zipfile, re

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
SRC = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx',
                   '0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx')
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', '0e7af_class_b')


def truncate_paragraphs(xml: str, keep_indices: list[int]) -> str:
    body_match = re.search(r'(<w:body[^>]*>)(.*?)(<w:sectPr[^>]*>.*?</w:sectPr>\s*</w:body>)', xml, re.DOTALL)
    if not body_match:
        return xml
    body_open = body_match.group(1)
    body_inner = body_match.group(2)
    body_tail = body_match.group(3)
    elements = re.findall(r'(?:<w:p\b[^>]*>.*?</w:p>|<w:tbl\b[^>]*>.*?</w:tbl>)', body_inner, re.DOTALL)
    kept = [elements[i-1] for i in keep_indices if 1 <= i <= len(elements)]
    new_body_inner = '\n'.join(kept) + '\n'
    return xml[:body_match.start()] + body_open + new_body_inner + body_tail + xml[body_match.end():]


def copy_docx_with_modified_xml(src_path: str, dst_path: str, new_doc_xml: str):
    os.makedirs(os.path.dirname(dst_path), exist_ok=True)
    with zipfile.ZipFile(src_path, 'r') as zin:
        with zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = new_doc_xml.encode('utf-8')
                zout.writestr(item, data)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    with zipfile.ZipFile(SRC, 'r') as z:
        with z.open('word/document.xml') as f:
            doc_xml = f.read().decode('utf-8')

    variants = [
        # V200: paragraph 207 only (the violator)
        ('CB_V200_para207_only', [207]),
        # V201: para 205-209 (with surrounding context for grouping)
        ('CB_V201_paras_205_to_209', list(range(205, 210))),
        # V202: para 200-215 (broader context)
        ('CB_V202_paras_200_to_215', list(range(200, 216))),
    ]

    for label, indices in variants:
        xml = truncate_paragraphs(doc_xml, indices)
        dst = os.path.join(OUT_DIR, f'{label}.docx')
        copy_docx_with_modified_xml(SRC, dst, xml)
        print(f'  wrote {label}.docx ({len(xml)} bytes XML, {len(indices)} paragraphs)')

    print('Done.')


if __name__ == '__main__':
    main()
