"""Replace unsupported fonts in test docx files with open/supported alternatives.

Ensures both Word and Oxi render with the same fonts, eliminating
font-related differences from the SSIM comparison.
"""
import zipfile, os, re, shutil, sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX_DIR = os.path.join(os.path.dirname(__file__), '..', '..',
                        'tools', 'golden-test', 'documents', 'docx')

# Font replacement map: unsupported → supported (with Oxi metrics)
FONT_MAP = {
    # Arial Unicode MS → MS Mincho (already aliased in code, make explicit in XML)
    'Arial Unicode MS': 'ＭＳ 明朝',
    # HG fonts → closest available
    'HGP創英角ｺﾞｼｯｸUB': 'ＭＳ ゴシック',
    'HGｺﾞｼｯｸM': 'ＭＳ ゴシック',
    'HG丸ｺﾞｼｯｸM-PRO': 'ＭＳ ゴシック',
    'HGﾃﾞｨｵﾐ': 'ＭＳ 明朝',
    # Chinese fonts → MS Mincho
    'PMingLiU': 'ＭＳ 明朝',
    # Courier → keep as-is (rare, ASCII only, width similar to MS Gothic)
    # Andale Sans UI → Arial
    'Andale Sans UI': 'Arial',
    # Georgia → Times New Roman
    'Georgia': 'Times New Roman',
    # Courier New → MS Gothic (monospace)
    'Courier New': 'ＭＳ ゴシック',
}

# XML parts to process
PARTS = ['word/document.xml', 'word/styles.xml', 'word/header1.xml',
         'word/header2.xml', 'word/header3.xml',
         'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml',
         'word/footnotes.xml', 'word/endnotes.xml']


def replace_fonts_in_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    """Replace font names in XML content. Returns (new_xml, replacement_count)."""
    text = xml_bytes.decode('utf-8')
    count = 0
    for old_font, new_font in FONT_MAP.items():
        if old_font in text:
            n = text.count(old_font)
            text = text.replace(old_font, new_font)
            count += n
    return text.encode('utf-8'), count


def process_docx(path: str) -> int:
    """Process a single docx file. Returns total replacement count."""
    total = 0
    tmp_path = path + '.tmp'

    with zipfile.ZipFile(path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in PARTS:
                    new_data, count = replace_fonts_in_xml(data)
                    if count > 0:
                        data = new_data
                        total += count
                zout.writestr(item, data)

    if total > 0:
        shutil.move(tmp_path, path)
    else:
        os.remove(tmp_path)

    return total


if __name__ == '__main__':
    docx_dir = os.path.abspath(DOCX_DIR)
    modified = 0

    for f in sorted(os.listdir(docx_dir)):
        if not f.endswith('.docx') or f.startswith('~$'):
            continue
        path = os.path.join(docx_dir, f)
        count = process_docx(path)
        if count > 0:
            print(f'  {f}: {count} replacements')
            modified += 1

    print(f'\nDone: {modified} files modified')
