# -*- coding: utf-8 -*-
"""Copy documents with every <w:lastRenderedPageBreak/> removed.

The marks are Word's own saved page breaks: an engine that reads them sees part
of the answer. Measuring on stripped copies against the unchanged Word truth
gives numbers no layout rule can have learned from the file.

    python tools/metrics/strip_lrpb.py <selection.json> <out_dir>

<selection.json> is a frozen blind selection ({type: [{"path": ...}, ...]}).
Copies land at <out_dir>/<type>/<sha>.docx; a per-document count of the marks
removed goes to <out_dir>/_stripped.json.
"""
import json, re, sys, zipfile
from pathlib import Path

MARK = re.compile(rb'<w:lastRenderedPageBreak\s*/>')


def strip(src: Path, dst: Path) -> int:
    n = 0
    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name.startswith('word/') and name.endswith('.xml'):
                data, k = MARK.subn(b'', data)
                n += k
            zout.writestr(name, data)
    return n


def main():
    sel = json.load(open(sys.argv[1], encoding='utf-8'))
    out = Path(sys.argv[2])
    counts = {}
    for kind, docs in sel.items():
        for d in docs:
            src = Path(d['path'])
            counts[f'{kind}__{src.stem}'] = strip(src, out / kind / src.name)
    json.dump(counts, open(out / '_stripped.json', 'w'), indent=1)
    print(f'{len(counts)} documents, {sum(counts.values())} marks removed, '
          f'{sum(1 for v in counts.values() if v)} documents had marks')


if __name__ == '__main__':
    main()
