"""Build long-paragraph + jc=both MS Mincho variant.

Strategy A retest: Earlier jc=both test (build_mincho_jc_both_variant.py)
used 観、「測 × 10 = 40 chars at 10.5pt = 420pt → fits one line → jc=both
has no justification work to do → may not have suppressed compression.

This variant: 観、「測 × 100 = 400 chars at 10.5pt → wraps to many lines
→ jc=both engages real justification work → tests "jc=both AND wrap-needed"
combinational hypothesis.

Expected outcomes:
- 、 compressed (~5.25pt) → jc=both is definitively NOT the gate even
  in wrap-needed context
- 、 full (~10.5pt) → jc=both + wrap-needed IS the gate (narrow Path B
  candidate; matches 0e7af situation)
"""
import os
import re
import sys
import zipfile

SRC = 'tools/metrics/mincho_adjacency_repro/MC_A_mincho.docx'
OUT = 'tools/metrics/jc_variants/MC_A_mincho_jc_both_long.docx'
NEW_TEXT = '観、「測' * 100  # 400 chars at 10.5pt → wraps to many lines


def transform(xml):
    # Step 1: replace text in the first paragraph
    first_t_replaced = [False]
    def t_replacer(match):
        if not first_t_replaced[0]:
            first_t_replaced[0] = True
            return f'<w:t xml:space="preserve">{NEW_TEXT}</w:t>'
        return '<w:t xml:space="preserve"></w:t>'
    xml = re.sub(r'<w:t[^>]*>[^<]*</w:t>', t_replacer, xml)

    # Step 2: add jc=both to the first paragraph
    def add_jc(body):
        ppr_match = re.search(r'<w:pPr>(.+?)</w:pPr>', body, flags=re.DOTALL)
        if ppr_match:
            inner = ppr_match.group(1)
            if '<w:jc' in inner:
                inner = re.sub(r'<w:jc\s+w:val="\w+"\s*/>', '<w:jc w:val="both"/>', inner)
            else:
                inner = '<w:jc w:val="both"/>' + inner
            return body[:ppr_match.start()] + f'<w:pPr>{inner}</w:pPr>' + body[ppr_match.end():]
        else:
            open_match = re.match(r'(<w:p\b[^>]*>)', body)
            if open_match:
                return open_match.group(1) + '<w:pPr><w:jc w:val="both"/></w:pPr>' + body[open_match.end():]
        return body

    para_pattern = re.compile(r'<w:p\b[^>]*>.+?</w:p>', flags=re.DOTALL)
    out_parts = []
    last_end = 0
    done = False
    for m in para_pattern.finditer(xml):
        out_parts.append(xml[last_end:m.start()])
        body = m.group(0)
        if not done and '<w:t' in body:
            body = add_jc(body)
            done = True
        out_parts.append(body)
        last_end = m.end()
    out_parts.append(xml[last_end:])
    return ''.join(out_parts)


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = transform(data.decode('utf-8')).encode('utf-8')
                zout.writestr(item, data)
    print(f'wrote {OUT}', flush=True)
    with zipfile.ZipFile(OUT) as z:
        d = z.read('word/document.xml').decode('utf-8')
    jcs = re.findall(r'<w:jc\s+w:val="(\w+)"', d)
    text_count = d.count('観')
    print(f'jc values: {jcs}, 観 count in text: {text_count}')


if __name__ == '__main__':
    sys.exit(main())
