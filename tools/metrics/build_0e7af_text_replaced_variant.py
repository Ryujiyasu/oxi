"""Replace 0e7af's first non-empty paragraph text with fixture pattern.

Tests whether yakumono compression in Word is text-shape-dependent or
doc-level-setting-dependent.

Procedure: take 0e7af.docx, find the first paragraph with `<w:t>`
content, replace ALL `<w:t>...</w:t>` content within that one paragraph
with `観、「測` × N (same as MC_A_mincho fixture). All other paragraphs
unchanged. All settings, styles, run/paragraph properties preserved.

If the modified paragraph compresses (`、` width ~5.25pt at 9pt font
i.e. ~4.5pt) → text-shape is the discriminator: real Japanese text in
0e7af doesn't compress, but synthetic `観、「測` does, even with
identical surrounding doc context.

If the modified paragraph does NOT compress (`、` width ~9.0pt) → some
doc-level setting in 0e7af suppresses compression universally. Need
to find that setting via binary diff vs MC_A_mincho.docx.
"""
import os
import re
import shutil
import sys
import zipfile

SRC = 'tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx'
OUT = 'tools/metrics/text_replaced_variants/0e7af_with_fixture_text_in_p1.docx'
PATTERN = '観、「測' * 10  # 40 chars matching MC_A_mincho text head


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    xml = data.decode('utf-8')
                    new_xml = replace_first_para_text(xml)
                    data = new_xml.encode('utf-8')
                zout.writestr(item, data)
    print(f'wrote {OUT}', flush=True)
    # Sanity
    with zipfile.ZipFile(OUT) as z:
        d = z.read('word/document.xml').decode('utf-8')
    # Count paragraphs and first paragraph text
    paras = re.findall(r'<w:p\b[^>]*>(.+?)</w:p>', d, flags=re.DOTALL)
    print(f'total paragraphs: {len(paras)}')
    for i, p in enumerate(paras):
        ts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p)
        joined = ''.join(ts)
        if joined.strip():
            print(f'first non-empty para idx {i}: text head = {joined[:30]!r}')
            break


def replace_first_para_text(xml):
    """Replace text in the first non-empty BODY paragraph (no pStyle, or body-style)."""
    para_pattern = re.compile(r'<w:p\b[^>]*>(.+?)</w:p>', flags=re.DOTALL)
    out_parts = []
    last_end = 0
    replaced = False
    for m in para_pattern.finditer(xml):
        out_parts.append(xml[last_end:m.start()])
        body = m.group(0)
        if not replaced:
            ts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', body)
            joined = ''.join(ts)
            if not joined.strip():
                out_parts.append(body)
                last_end = m.end()
                continue
            # Skip paragraphs with pStyle (likely headings or special styles).
            # We want a plain body paragraph that inherits default body font.
            has_pstyle = bool(re.search(r'<w:pStyle\b', body))
            # Also skip if any run has explicit font size (heading-likely)
            has_run_size = bool(re.search(r'<w:sz\s+w:val="(\d+)"', body))
            if has_pstyle or has_run_size:
                # Keep moving forward
                out_parts.append(body)
                last_end = m.end()
                continue
            # This is a plain body paragraph — replace text
            first_t_replaced = [False]
            def t_replacer(match):
                if not first_t_replaced[0]:
                    first_t_replaced[0] = True
                    return f'<w:t xml:space="preserve">{PATTERN}</w:t>'
                else:
                    return '<w:t xml:space="preserve"></w:t>'
            body = re.sub(r'<w:t[^>]*>[^<]*</w:t>', t_replacer, body)
            replaced = True
        out_parts.append(body)
        last_end = m.end()
    out_parts.append(xml[last_end:])
    if not replaced:
        print('WARN: no body paragraph found, did not replace anything', file=sys.stderr)
    return ''.join(out_parts)


if __name__ == '__main__':
    sys.exit(main())
