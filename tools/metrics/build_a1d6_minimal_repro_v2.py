"""a1d6 から T5 R1-R3 だけ残した minimal docx 生成 (regex ベース、XML 構造保持).

ET.tostring は docx の namespace prefix を壊すので使えない。
代わりに source の document.xml を文字列のまま読み込み、T5 の特定 <w:tr> 要素を
regex/string slice で削除する。
"""
import sys, os, zipfile, re
sys.stdout.reconfigure(encoding='utf-8')

SRC = 'tools/golden-test/documents/docx/a1d6e4efa2e7_tokumei_08_01-4.docx'
OUT_V2 = 'c:/tmp/minimal_a1d6_v2.docx'

def find_table_ranges(xml):
    """各 <w:tbl> ... </w:tbl> の (start, end) を返す.
    nested table 対応のため stack で深さ管理 (ただし a1d6 にはなし)."""
    tables = []
    stack = []
    i = 0
    while i < len(xml):
        m_open = re.search(r'<w:tbl(?:\s|>)', xml[i:])
        m_close = re.search(r'</w:tbl>', xml[i:])
        if not m_open and not m_close: break
        next_open = i + m_open.start() if m_open else float('inf')
        next_close = i + m_close.start() if m_close else float('inf')
        if next_open < next_close:
            stack.append(next_open)
            i = next_open + 1
        else:
            start = stack.pop()
            end = next_close + len('</w:tbl>')
            if not stack:  # top-level table
                tables.append((start, end))
            i = end
    return tables

def find_tr_ranges(tbl_xml):
    """テーブル内の <w:tr> ... </w:tr> ranges (offsets relative to tbl_xml)."""
    trs = []
    i = 0
    while True:
        m = re.search(r'<w:tr(?:\s|>)', tbl_xml[i:])
        if not m: break
        start = i + m.start()
        m2 = re.search(r'</w:tr>', tbl_xml[start:])
        if not m2: break
        end = start + m2.start() + len('</w:tr>')
        trs.append((start, end))
        i = end
    return trs

def main():
    with zipfile.ZipFile(SRC) as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    doc_xml = contents['word/document.xml'].decode('utf-8')
    tables = find_table_ranges(doc_xml)
    print(f'Found {len(tables)} top-level tables')

    # T5 = tables[4] (0-based)
    t5_start, t5_end = tables[4]
    t5_xml = doc_xml[t5_start:t5_end]
    trs = find_tr_ranges(t5_xml)
    print(f'T5 has {len(trs)} tr elements')

    # R1-R3 (indices 0-2) は残し、R4-R30 (indices 3-29) を削除
    # 削除は後ろから (オフセット保持)
    keep_trs = trs[:3]
    new_t5_inner = t5_xml[:trs[2][1]]  # R1-R3 終わりまで
    # tbl の閉じタグまでの末尾を保持
    tail = t5_xml[trs[-1][1]:]  # 最後の tr の後 (もし tail にコンテンツがあれば残す)
    new_t5_xml = new_t5_inner + tail
    print(f'v2 T5 xml length: {len(new_t5_xml)} (was {len(t5_xml)})')

    # 元 doc xml に書き戻す
    new_doc_xml = doc_xml[:t5_start] + new_t5_xml + doc_xml[t5_end:]
    contents['word/document.xml'] = new_doc_xml.encode('utf-8')

    with zipfile.ZipFile(OUT_V2, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents[n])
    print(f'Wrote {OUT_V2}')

if __name__ == '__main__':
    main()
