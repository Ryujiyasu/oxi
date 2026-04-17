"""Strip d77a by deleting paragraphs NOT containing target text."""
import os, sys, time, zipfile, re
import win32com.client

TMP = os.path.abspath("pipeline_data/_ctx_tmp")
os.makedirs(TMP, exist_ok=True)
SRC = os.path.abspath(r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")

KEEP_TEXT = "公共データ利用規約（第1.0版"  # matches idx=10

def smart_strip(src, out):
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == 'word/document.xml':
                    xml = data.decode('utf-8')
                    # Split by </w:p> to isolate paragraphs
                    parts = xml.split('</w:p>')
                    # For each part, check if it contains the target
                    # parts[-1] is the trailing XML after last </w:p>
                    kept_parts = []
                    found = False
                    for i, part in enumerate(parts[:-1]):
                        # part + '</w:p>' is one paragraph block
                        if KEEP_TEXT in part:
                            kept_parts.append(part + '</w:p>')
                            found = True
                            break
                    # Find the start of first paragraph in original (to keep header before first <w:p>)
                    m_body = re.search(r'(<w:body>)', xml)
                    body_start_idx = m_body.end() if m_body else 0
                    # Get everything before first <w:p>
                    m_firstp = re.search(r'<w:p\b', xml)
                    header_end_idx = m_firstp.start() if m_firstp else body_start_idx
                    header = xml[:header_end_idx]
                    # sectPr: find in original xml
                    m_sect = re.search(r'<w:sectPr\b[^/]*/>|<w:sectPr\b[^>]*>.*?</w:sectPr>', xml, re.DOTALL)
                    sect = m_sect.group(0) if m_sect else ''
                    # Footer
                    m_body_end = re.search(r'</w:body>.*$', xml, re.DOTALL)
                    footer = m_body_end.group(0) if m_body_end else '</w:body></w:document>'
                    new_xml = header + ''.join(kept_parts) + sect + footer
                    data = new_xml.encode('utf-8')
                    if not found:
                        print(f"WARN: KEEP_TEXT not found in paragraphs")
                zout.writestr(item, data)

def measure(path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
        for p in doc.Paragraphs:
            text = p.Range.Text
            if '（' not in text: continue
            rng = p.Range
            for ci in range(1, rng.Characters.Count + 1):
                c = rng.Characters(ci)
                if c.Text == '（':
                    x1 = c.Information(5); y1 = c.Information(6)
                    try:
                        nxt = rng.Characters(ci + 1)
                        x2 = nxt.Information(5); y2 = nxt.Information(6)
                        if abs(y1 - y2) > 2: continue
                        doc.Close(False)
                        return round(x2 - x1, 2), text[:50]
                    except: pass
        doc.Close(False)
        return None, None
    finally:
        word.Quit()

out = os.path.join(TMP, "d77a_v2_idx10.docx")
smart_strip(SRC, out)
print(f"stripped size: {os.path.getsize(out)} bytes")
adv, text = measure(out)
print(f"advance={adv}  text={text!r}")
