"""a1d6 T5 から特定の subset を残した minimal repro docx を生成する。

T5 (30-row vMerge テーブル) の中で APPLY 7 行 (R4, R11, R12, R13, R24, R25, R26)
の挙動を切り分けるための実験用 variants:

  v1: R27-R30 を削除 (T5 末尾の SUPPRESS rows と次セクション独立行を取り除く)
      → R24/R25 が APPLY のままか?
      仮説 A: 次セクション独立行が R23 セクションの APPLY を誘発 → R24/R25 が SUPPRESS に変わる
      仮説 B: R23-R26 構造内に APPLY 要因がある → R24/R25 は APPLY のまま

入力: tools/golden-test/documents/docx/a1d6e4efa2e7_tokumei_08_01-4.docx
出力: c:/tmp/minimal_a1d6_v1.docx
"""
import sys, os, zipfile, shutil, re
from io import BytesIO
import xml.etree.ElementTree as ET

SRC = 'tools/golden-test/documents/docx/a1d6e4efa2e7_tokumei_08_01-4.docx'
OUT_V1 = 'c:/tmp/minimal_a1d6_v1.docx'
OUT_V2 = 'c:/tmp/minimal_a1d6_v2.docx'  # T5 = R1-R3 only (test row pitch hypothesis)

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def main():
    # ET だと namespace prefix が ns0 などに書き換わって Word が読めないことがある
    # 代わりに lxml が安全だが、ここでは ET の register_namespace で対応
    ET.register_namespace('', NS['w'])
    for prefix, uri in {
        'w': NS['w'],
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'v': 'urn:schemas-microsoft-com:vml',
        'o': 'urn:schemas-microsoft-com:office:office',
        'w10': 'urn:schemas-microsoft-com:office:word',
    }.items():
        ET.register_namespace(prefix, uri)

    # Read source docx
    with zipfile.ZipFile(SRC) as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    # === v1: T5 R27-R30 を削除 ===
    raw_doc = contents['word/document.xml'].decode('utf-8')
    tree = ET.fromstring(raw_doc)
    tables = tree.findall('.//w:tbl', NS)
    T5 = tables[4]
    rows = T5.findall('w:tr', NS)
    print(f'Original T5 has {len(rows)} rows')

    for ri in range(29, 25, -1):
        if ri < len(rows):
            T5.remove(rows[ri])
    new_rows = T5.findall('w:tr', NS)
    print(f'v1 T5 has {len(new_rows)} rows after deletion')

    new_xml = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
    contents_v1 = dict(contents)
    contents_v1['word/document.xml'] = new_xml

    os.makedirs(os.path.dirname(OUT_V1), exist_ok=True)
    with zipfile.ZipFile(OUT_V1, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents_v1[n])
    print(f'Wrote {OUT_V1}')

    # === v2: T5 R1-R3 だけ残す (仮説 5 検証用) ===
    raw_doc2 = contents['word/document.xml'].decode('utf-8')
    tree2 = ET.fromstring(raw_doc2)
    tables2 = tree2.findall('.//w:tbl', NS)
    T5_v2 = tables2[4]
    rows2 = T5_v2.findall('w:tr', NS)
    # R4-R30 削除 (3 行だけ残す)
    for ri in range(len(rows2)-1, 2, -1):
        T5_v2.remove(rows2[ri])
    new_rows2 = T5_v2.findall('w:tr', NS)
    print(f'v2 T5 has {len(new_rows2)} rows (expect 3)')

    new_xml2 = ET.tostring(tree2, encoding='utf-8', xml_declaration=True)
    contents_v2 = dict(contents)
    contents_v2['word/document.xml'] = new_xml2

    with zipfile.ZipFile(OUT_V2, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents_v2[n])
    print(f'Wrote {OUT_V2}')

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    main()
