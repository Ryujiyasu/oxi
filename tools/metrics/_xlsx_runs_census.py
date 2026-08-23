# How many workbooks hold a wrapped cell whose text is dressed in pieces.
import re, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
DOCS = Path(r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\xlsx")
sys.stdout.reconfigure(encoding="utf-8")
books, cells_total = 0, 0
worst = []
for book in sorted(DOCS.glob("*.xlsx")):
    try:
        z = zipfile.ZipFile(book)
        shared = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    except Exception:
        continue
    # Which shared strings hold more than one dressed run.
    dressed = set()
    for index, si in enumerate(re.findall(r"<si>.*?</si>", shared, re.S)):
        if si.count("<r>") > 1 and "<rPr>" in si:
            dressed.add(index)
    if not dressed:
        continue
    styles = z.read("xl/styles.xml").decode("utf-8", "replace")
    xfs = re.search(r'<cellXfs count="\d+">(.*?)</cellXfs>', styles, re.S)
    items = re.findall(r"<xf [^>]*/>|<xf .*?</xf>", xfs.group(1), re.S) if xfs else []
    wraps = {i for i, x in enumerate(items) if 'wrapText="1"' in x}
    held = 0
    for part in z.namelist():
        if not re.match(r"xl/worksheets/sheet\d+\.xml", part):
            continue
        sheet = z.read(part).decode("utf-8", "replace")
        for m in re.finditer(r'<c r="[A-Z]+\d+"([^>]*)>\s*<v>(\d+)</v>', sheet):
            attrs, value = m.groups()
            if int(value) not in dressed or 't="s"' not in attrs:
                continue
            s = re.search(r's="(\d+)"', attrs)
            if s and int(s.group(1)) in wraps:
                held += 1
    if held:
        books += 1
        cells_total += held
        worst.append((held, book.stem))
worst.sort(reverse=True)
print(f"{books} workbooks hold {cells_total} wrapped cells dressed in pieces")
for held, stem in worst[:12]:
    print(f"  {held:>4}  {stem}")
