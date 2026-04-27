"""Test whether jc=both (justify) suppresses yakumono compression.

After 4 single-axis hypotheses REFUTED (compat, kern, font, size), and
text-injection into 0e7af STILL not compressing, the suspect narrowed
to doc-level pPrDefault. 0e7af has <w:jc w:val="both"/> in pPrDefault;
MC_A_mincho has no styles.xml at all (Word default = jc=left, which
compresses).

This script takes MC_A_mincho.docx and adds jc=both to its single
paragraph's pPr. Run + measure to test:
- if `、` becomes 10.5pt (full) → jc=both is the gate
- if `、` stays 5.25pt (compressed) → jc=both is not the gate (look elsewhere)
"""
import os
import re
import sys
import zipfile

SRC = 'tools/metrics/mincho_adjacency_repro/MC_A_mincho.docx'
OUT = 'tools/metrics/jc_variants/MC_A_mincho_jc_both.docx'


def add_jc_both(xml):
    # Find the first <w:p> element and inject <w:pPr><w:jc w:val="both"/></w:pPr>
    # right after the opening tag. Skip if pPr already exists, in which case
    # add jc inside it.
    def add_jc(body):
        # Check for existing pPr
        ppr_match = re.search(r'<w:pPr>(.+?)</w:pPr>', body, flags=re.DOTALL)
        if ppr_match:
            inner = ppr_match.group(1)
            if '<w:jc' in inner:
                # Replace existing jc
                inner = re.sub(r'<w:jc\s+w:val="\w+"\s*/>', '<w:jc w:val="both"/>', inner)
            else:
                inner = '<w:jc w:val="both"/>' + inner
            new_ppr = f'<w:pPr>{inner}</w:pPr>'
            body = body[:ppr_match.start()] + new_ppr + body[ppr_match.end():]
        else:
            # Insert pPr right after the opening <w:p ...> tag
            open_match = re.match(r'(<w:p\b[^>]*>)', body)
            if open_match:
                body = open_match.group(1) + '<w:pPr><w:jc w:val="both"/></w:pPr>' + body[open_match.end():]
        return body

    # Apply to first <w:p>...</w:p> only
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
                    data = add_jc_both(data.decode('utf-8')).encode('utf-8')
                zout.writestr(item, data)
    print(f'wrote {OUT}', flush=True)
    # Sanity
    with zipfile.ZipFile(OUT) as z:
        d = z.read('word/document.xml').decode('utf-8')
    jcs = re.findall(r'<w:jc\s+w:val="(\w+)"', d)
    print(f'jc values: {jcs}')


if __name__ == '__main__':
    sys.exit(main())
