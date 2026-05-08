"""Generate V116-V120 repros that capture db9ca page-break drift event.

Day 31 part 9 falsified isolated single-paragraph hypothesis. The +18pt
drift in db9ca emerges from PAGE BREAK + multi-paragraph context.

Strategy: copy db9ca docx but truncate paragraphs after a specific cutoff,
preserving styles/settings/sections so rendering context matches.

Variants:
  V116: keep paragraphs 1-22 (covers page break + drift event paragraphs 17-20)
  V117: keep paragraphs 1-25 (extends past drift jump for cum check)
  V118: keep paragraphs 1-30 (covers 2nd +18pt jump location at i=31)
  V119: keep paragraphs 1-15 (page 1 only, no page break = control)
  V120: keep paragraphs 14-22 (drift event in isolation, no preceding context)
"""
from __future__ import annotations
import os, zipfile, re, shutil, io

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
SRC = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx',
                   'db9ca18368cd_20241122_resource_open_data_01.docx')
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'db9ca_section')


def truncate_paragraphs(xml: str, keep_first_n: int | None = None,
                        keep_range: tuple[int, int] | None = None) -> str:
    """Keep paragraphs by 1-indexed position. If keep_first_n=N, keep 1..N.
    If keep_range=(start,end), keep start..end (1-indexed inclusive).
    """
    # Find body content (between <w:body...> and <w:sectPr...>) — split by paragraphs
    # Note: w:body and w:sectPr may have attributes (rsid*)
    body_match = re.search(r'(<w:body[^>]*>)(.*?)(<w:sectPr[^>]*>.*?</w:sectPr>\s*</w:body>)', xml, re.DOTALL)
    if not body_match:
        return xml
    body_open = body_match.group(1)
    body_inner = body_match.group(2)
    body_tail = body_match.group(3)

    # Extract paragraphs (and tables) preserving order
    elements = re.findall(r'(?:<w:p\b[^>]*>.*?</w:p>|<w:tbl\b[^>]*>.*?</w:tbl>)', body_inner, re.DOTALL)

    # Apply selection
    if keep_first_n is not None:
        kept = elements[:keep_first_n]
    elif keep_range is not None:
        s, e = keep_range
        kept = elements[s-1:e]
    else:
        kept = elements

    new_body_inner = '\n'.join(kept) + '\n'
    return xml[:body_match.start()] + body_open + new_body_inner + body_tail + xml[body_match.end():]


def copy_docx_with_modified_xml(src_path: str, dst_path: str, new_doc_xml: str):
    """Copy zip preserving all entries, replacing word/document.xml."""
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
        ('DS_V116_paras_1_to_22', truncate_paragraphs(doc_xml, keep_first_n=22)),
        ('DS_V117_paras_1_to_25', truncate_paragraphs(doc_xml, keep_first_n=25)),
        ('DS_V118_paras_1_to_30', truncate_paragraphs(doc_xml, keep_first_n=30)),
        ('DS_V119_paras_1_to_15', truncate_paragraphs(doc_xml, keep_first_n=15)),
        ('DS_V120_paras_14_to_22', truncate_paragraphs(doc_xml, keep_range=(14, 22))),
        # Day 31 part 10 — current Oxi shows +19pt drift jump at i=43 (not i=20)
        ('DS_V122_paras_40_to_46', truncate_paragraphs(doc_xml, keep_range=(40, 46))),
        ('DS_V123_paras_1_to_45', truncate_paragraphs(doc_xml, keep_first_n=45)),
        ('DS_V124_paras_30_to_46', truncate_paragraphs(doc_xml, keep_range=(30, 46))),
        # Day 31 part 15 — wrap-width bug isolation: extreme divergence paragraphs
        ('DS_V125_para11_only', truncate_paragraphs(doc_xml, keep_range=(11, 11))),
        ('DS_V126_para15_only', truncate_paragraphs(doc_xml, keep_range=(15, 15))),
        ('DS_V127_para25_only', truncate_paragraphs(doc_xml, keep_range=(25, 25))),
    ]

    for label, xml in variants:
        dst = os.path.join(OUT_DIR, f'{label}.docx')
        copy_docx_with_modified_xml(SRC, dst, xml)
        print(f'  wrote {label}.docx ({len(xml)} bytes XML)')

    print('Done.')


if __name__ == '__main__':
    main()
